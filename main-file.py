import os
import sys
import argparse
from excel_analyzer import ExcelAnalyzer
from embedders import BedrockEmbedder
from database import VectorDatabase

def process_euda_file(file_path):
    """Process a single EUDA Excel file and store its embeddings"""
    print(f"Processing EUDA file: {file_path}")
    
    # Initialize components
    analyzer = ExcelAnalyzer()
    embedder = BedrockEmbedder()
    db = VectorDatabase()
    
    try:
        # Connect to database
        db.connect()
        
        # Load and analyze Excel file
        if not analyzer.load_file(file_path):
            print(f"Failed to load file: {file_path}")
            return False
        
        # Analyze file and extract metadata
        metadata = analyzer.analyze_file()
        if not metadata:
            print(f"Failed to analyze file: {file_path}")
            return False
        
        # Store file metadata in database
        file_name = os.path.basename(file_path)
        file_size_kb = metadata["file_size_kb"]
        sheet_count = metadata["sheet_count"]
        
        euda_id = db.store_euda_metadata(file_name, file_path, file_size_kb, sheet_count, metadata)
        if not euda_id:
            print(f"Failed to store metadata for file: {file_path}")
            return False
        
        # Extract text chunks for embedding
        text_chunks = analyzer.extract_text_for_embedding()
        
        # Process each text chunk
        for chunk in text_chunks:
            # Generate embedding for the text chunk
            embedding = embedder.get_text_embedding(chunk["content"])
            
            if embedding:
                # Store embedding in database
                db.store_embedding(
                    euda_id=euda_id,
                    content_type=chunk["type"],
                    content_text=chunk["content"],
                    embedding=embedding
                )
            else:
                print(f"Failed to generate embedding for chunk type: {chunk['type']}")
        
        print(f"Successfully processed EUDA file: {file_path}")
        return True
    
    except Exception as e:
        print(f"Error processing EUDA file: {str(e)}")
        return False
    
    finally:
        # Clean up
        analyzer.close()
        db.close()

def main():
    """Main entry point of the application"""
    parser = argparse.ArgumentParser(description="EUDA Excel Analyzer and Vector Embedder")
    parser.add_argument("--file", "-f", help="Path to Excel EUDA file to analyze")
    parser.add_argument("--directory", "-d", help="Directory containing Excel EUDA files to analyze")
    
    args = parser.parse_args()
    
    if not args.file and not args.directory:
        print("Error: Either --file or --directory must be specified.")
        parser.print_help()
        return 1
    
    if args.file:
        # Process single file
        if not os.path.isfile(args.file):
            print(f"Error: File not found: {args.file}")
            return 1
        
        success = process_euda_file(args.file)
        if not success:
            return 1
    
    if args.directory:
        # Process all Excel files in directory
        if not os.path.isdir(args.directory):
            print(f"Error: Directory not found: {args.directory}")
            return 1
        
        excel_extensions = ['.xlsx', '.xls', '.xlsm', '.xlsb']
        files_processed = 0
        files_failed = 0
        
        for file_name in os.listdir(args.directory):
            file_path = os.path.join(args.directory, file_name)
            
            # Check if file is an Excel file
            if os.path.isfile(file_path) and any(file_name.lower().endswith(ext) for ext in excel_extensions):
                success = process_euda_file(file_path)
                
                if success:
                    files_processed += 1
                else:
                    files_failed += 1
        
        print(f"Summary: {files_processed} files processed successfully, {files_failed} files failed.")
    
    return 0

if __name__ == "__main__":
    sys.exit(main())
