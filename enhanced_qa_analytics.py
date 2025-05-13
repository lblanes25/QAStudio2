#!/usr/bin/env python3
"""
Main script for the Enhanced QA Analytics Automation Framework.
"""

import os
import sys
import argparse
import logging
from tkinter import Tk

from logging_config import setup_logging
from enhanced_qa_analytics_app import EnhancedQAAnalyticsApp

logger = setup_logging()


def create_directories():
    """Create required directories if they don't exist"""
    required_dirs = [
        "configs",  # Configuration files
        "output",  # Report output
        "ref_data",  # Reference data
        "test_data",  # Test data
        "logs"  # Log files
    ]

    for directory in required_dirs:
        if not os.path.exists(directory):
            os.makedirs(directory)
            logger.info(f"Created directory: {directory}")


def run_gui_mode():
    """Run the application in GUI mode"""
    root = Tk()
    app = EnhancedQAAnalyticsApp(root)
    root.mainloop()


def run_cli_mode(args):
    """
    Run the application in command-line mode

    Args:
        args: Command-line arguments
    """
    try:
        from config_manager import ConfigManager
        from enhanced_data_processor import EnhancedDataProcessor
        from enhanced_report_generator import EnhancedReportGenerator

        # Validate analytics ID
        if not args.analytic_id:
            logger.error("No analytics ID specified")
            return 1

        # Validate source file
        if not args.source_file or not os.path.exists(args.source_file):
            logger.error(f"Source file not found: {args.source_file}")
            return 1

        # Ensure output directory exists
        output_dir = args.output_dir or "output"
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # Load configuration
        config_manager = ConfigManager()

        try:
            config = config_manager.get_config(args.analytic_id)
        except ValueError as e:
            logger.error(f"Error loading config: {e}")
            return 1

        # Process data
        logger.info(f"Processing QA-ID {args.analytic_id} with source file {args.source_file}")
        processor = EnhancedDataProcessor(config)
        success, message = processor.process_data(args.source_file)

        if not success:
            logger.error(f"Processing failed: {message}")
            return 1

        logger.info(f"Processing completed: {message}")

        # Generate reports
        logger.info("Generating reports...")
        report_generator = EnhancedReportGenerator(config, processor.results)

        # Generate main report
        import datetime
        main_report_path = os.path.join(
            output_dir,
            f"QA_{args.analytic_id}_Main_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        main_report = report_generator.generate_main_report(main_report_path, args.source_file)

        if not main_report:
            logger.error("Failed to generate main report")
            return 1

        # Generate individual reports if requested
        individual_reports = []
        if args.individual_reports:
            individual_reports = report_generator.generate_individual_reports()

        # Show completion message
        report_count = 1 + len(individual_reports)
        logger.info(f"Processing complete. Generated {report_count} reports.")
        logger.info(f"Main report: {main_report}")

        for report in individual_reports:
            logger.info(f"Individual report: {report}")

        return 0

    except Exception as e:
        logger.error(f"Error in CLI mode: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return 1


def parse_arguments():
    """Parse command-line arguments"""
    parser = argparse.ArgumentParser(description="Enhanced QA Analytics Automation")

    # Common arguments
    parser.add_argument("--gui", action="store_true", help="Run in GUI mode (default if no arguments provided)")

    # CLI mode arguments
    parser.add_argument("-a", "--analytic-id", help="Analytic ID to run")
    parser.add_argument("-s", "--source-file", help="Source data file path")
    parser.add_argument("-o", "--output-dir", help="Output directory for reports")
    parser.add_argument("-i", "--individual-reports", action="store_true", help="Generate individual reports")

    args = parser.parse_args()

    # Default to GUI mode if no specific CLI arguments provided
    if len(sys.argv) == 1 or (args.gui and not (args.analytic_id or args.source_file)):
        args.gui = True

    return args


def main():
    """Main entry point"""
    # Create required directories
    create_directories()

    # Parse command-line arguments
    args = parse_arguments()

    # Run in appropriate mode
    if args.gui:
        logger.info("Starting Enhanced QA Analytics in GUI mode")
        run_gui_mode()
        return 0
    else:
        logger.info("Starting Enhanced QA Analytics in CLI mode")
        return run_cli_mode(args)


if __name__ == "__main__":
    sys.exit(main())