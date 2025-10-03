#!/usr/bin/env python3
"""
Utility for merging Excel tables

Algorithm:
1. Analyzes tables from the data folder
2. Finds header rows (row 4 in all files)
3. Creates unified header from all unique fields
4. Merges data into single table
5. Adds column with source filename
"""

import pandas as pd
import os
import logging
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Set

class SimplifiedTableFusion:
    """Simplified version for merging tables"""
    
    def __init__(self, data_dir: str = "data", result_dir: str = "result"):
        """Initialization"""
        self.data_dir = Path(data_dir)
        self.result_dir = Path(result_dir)
        
        # Logging configuration
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        self.logger = logging.getLogger(__name__)
        
        # Create results folder
        self.result_dir.mkdir(exist_ok=True)
    
    def get_xlsx_files(self) -> List[Path]:
        """Get all XLSX files from data folder and subfolders, sorted alphabetically"""
        if not self.data_dir.exists():
            raise FileNotFoundError(f"Data folder '{self.data_dir}' not found")
        
        # Search recursively in all subfolders
        xlsx_files = list(self.data_dir.rglob("*.xlsx"))
        
        # Sort files alphabetically by full path for consistent ordering
        xlsx_files.sort(key=lambda x: str(x).lower())
        
        self.logger.info(f"Found {len(xlsx_files)} XLSX files in data folder and subfolders")
        for i, file_path in enumerate(xlsx_files, 1):
            # Show relative path from data_dir for better readability
            relative_path = file_path.relative_to(self.data_dir)
            self.logger.info(f"  {i}. {relative_path}")
        
        return xlsx_files
    
    def find_header_row(self, df: pd.DataFrame) -> int:
        """
        Find header row
        
        Args:
            df: DataFrame for analysis
            
        Returns:
            Header row number or -1 if not found
        """
        # Check first 10 rows
        for i in range(min(10, len(df))):
            row = df.iloc[i]
            # Count non-empty values
            non_null_count = row.notna().sum()
            
            # If row has minimum 5 non-empty values, it might be a header
            if non_null_count >= 5:
                # Check for typical headers
                row_str = ' '.join(row.astype(str).str.lower())
                if any(header in row_str for header in ['title', 'composer', 'artist', 'album']):
                    self.logger.info(f"Found header row: {i}")
                    return i
        
        self.logger.warning("Header row not found")
        return -1
    
    def read_xlsx_with_headers(self, file_path: Path) -> pd.DataFrame:
        """
        Reads XLSX file with correct headers
        
        Args:
            file_path: File path
            
        Returns:
            DataFrame with data or None on error
        """
        try:
            # Read file without headers
            df_raw = pd.read_excel(file_path, header=None)
            
            # Find header row
            header_row = self.find_header_row(df_raw)
            
            if header_row == -1:
                self.logger.warning(f"Failed to find headers in {file_path.name}")
                return None
            
            # Re-read with correct headers
            df = pd.read_excel(file_path, header=header_row)
            
            # Remove empty rows
            df = df.dropna(how='all')
            
            # Add column with relative path from data directory
            relative_path = file_path.relative_to(self.data_dir)
            df['source_file'] = str(relative_path.with_suffix(''))
            
            self.logger.info(f"Read {file_path.name}: {len(df)} rows, {len(df.columns)} columns")
            return df
            
        except Exception as e:
            self.logger.error(f"Error reading {file_path.name}: {e}")
            return None
    
    def get_all_unique_columns(self, dataframes: List[pd.DataFrame]) -> List[str]:
        """
        Get all unique headers from all tables
        Preserves header order from first file, adds new ones at the end
        
        Args:
            dataframes: List of DataFrames
            
        Returns:
            List of unique headers in correct order
        """
        if not dataframes:
            return ['source_file']
        
        # Take headers from first file as base order
        first_file_columns = [col for col in dataframes[0].columns if col != 'source_file']
        
        # Collect all unique headers from other files
        all_columns = set(first_file_columns)
        
        for df in dataframes[1:]:
            columns = [col for col in df.columns if col != 'source_file']
            all_columns.update(columns)
        
        # Create final list: first order from first file
        ordered_columns = first_file_columns.copy()
        
        # Then add new headers from other files
        new_columns = [col for col in all_columns if col not in first_file_columns]
        ordered_columns.extend(sorted(new_columns))  # Sort only new ones for consistency
        
        # Add source_file at the end
        ordered_columns.append('source_file')
        
        self.logger.info(f"Header order: {len(first_file_columns)} from first file + {len(new_columns)} new + source_file")
        return ordered_columns
    
    def merge_dataframes(self, dataframes: List[pd.DataFrame]) -> pd.DataFrame:
        """
        Merges DataFrames into one with common headers
        
        Args:
            dataframes: List of DataFrames to merge
            
        Returns:
            Merged DataFrame
        """
        if not dataframes:
            raise ValueError("No data to merge")
        
        # Get all unique headers
        all_columns = self.get_all_unique_columns(dataframes)
        
        # Normalize all DataFrames to unified structure
        normalized_dfs = []
        for df in dataframes:
            # Create new DataFrame with complete set of columns
            normalized_df = pd.DataFrame(columns=all_columns)
            
            # Copy existing data
            for col in df.columns:
                if col in all_columns:
                    normalized_df[col] = df[col].values
            
            normalized_dfs.append(normalized_df)
        
        # Merge all DataFrames
        result_df = pd.concat(normalized_dfs, ignore_index=True)
        
        # Sort by source file
        result_df = result_df.sort_values('source_file', ignore_index=True)
        
        self.logger.info(f"Merged into table: {len(result_df)} rows, {len(result_df.columns)} columns")
        return result_df
    
    def save_result(self, df: pd.DataFrame) -> Path:
        """
        Saves result to Excel file
        
        Args:
            df: DataFrame to save
            
        Returns:
            Path to saved file
        """
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"{timestamp}.xlsx"
        output_path = self.result_dir / filename
        
        try:
            # Save with headers
            df.to_excel(output_path, index=False, engine='openpyxl')
            self.logger.info(f"Result saved: {output_path}")
            return output_path
            
        except Exception as e:
            self.logger.error(f"Save error: {e}")
            raise
    
    def show_summary(self, df: pd.DataFrame):
        """Shows summary of results"""
        if 'source_file' not in df.columns:
            return
        
        file_stats = df['source_file'].value_counts().sort_index()
        
        print(f"\n{'='*60}")
        print("📊 TABLE MERGING SUMMARY")
        print(f"{'='*60}")
        print(f"Total files processed: {len(file_stats)}")
        print(f"Total data rows: {len(df)}")
        print(f"Number of columns: {len(df.columns)}")
        print(f"\nRows by file:")
        
        for filename, count in file_stats.items():
            print(f"  • {filename}: {count} rows")
        
        print(f"\nColumn headers:")
        for i, col in enumerate(df.columns, 1):
            print(f"  {i:2d}. {col}")
        
        # Show folder structure info
        print(f"\n📁 Folder structure:")
        print(f"  Data folder: {self.data_dir}")
        print(f"  Result folder: {self.result_dir}")
    
    def run(self) -> Path:
        """
        Main execution method
        
        Returns:
            Path to result file
        """
        self.logger.info("Starting simplified table merging algorithm...")
        
        # Get list of files
        xlsx_files = self.get_xlsx_files()
        
        if not xlsx_files:
            raise ValueError("No XLSX files found in data folder")
        
        # Read all files
        dataframes = []
        for file_path in xlsx_files:
            df = self.read_xlsx_with_headers(file_path)
            if df is not None and not df.empty:
                dataframes.append(df)
        
        if not dataframes:
            raise ValueError("Failed to read any file with data")
        
        # Merge data
        merged_df = self.merge_dataframes(dataframes)
        
        # Save result
        output_path = self.save_result(merged_df)
        
        # Show summary
        self.show_summary(merged_df)
        
        self.logger.info("Table merging completed successfully!")
        return output_path


def main():
    """Main function"""
    try:
        fusion = SimplifiedTableFusion()
        output_path = fusion.run()
        
        print(f"\n✅ Success! Merged table saved: {output_path}")
        print(f"📁 Check the result folder to view the output.")
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
        logging.error(f"Execution error: {e}")
        return 1
    
    return 0


if __name__ == "__main__":
    exit(main())
