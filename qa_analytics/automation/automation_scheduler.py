# First, let me add the necessary imports at the top of the file
import os
import time
import datetime
import threading
import logging
import yaml
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from typing import Dict, List, Optional, Union, Tuple
import fnmatch

# Import schedule library
import schedule

# Import Excel processing components
from qa_analytics.core.excel_engine import ExcelFormulaProcessor, ensure_excel_closed
from qa_analytics.core.excel_utils import convert_excel_errors_to_none

# Set up logging
logger = logging.getLogger("qa_analytics")


class AutomationScheduler:
    """Manages scheduling and automated execution of QA analytics"""

    def __init__(self, config_manager, data_processor_class, report_generator_class):
        """
        Initialize the automation scheduler

        Args:
            config_manager: ConfigManager instance for loading/saving configs
            data_processor_class: EnhancedDataProcessor class (not instance)
            report_generator_class: EnhancedReportGenerator class (not instance)
        """
        self.config_manager = config_manager
        self.data_processor_class = data_processor_class
        self.report_generator_class = report_generator_class

        # Initialize scheduler
        self.scheduler = schedule
        self.running = False
        self.scheduler_thread = None

        # Excel processing tracking
        self.excel_processors = {}  # job_id -> ExcelFormulaProcessor

        # Load scheduler configuration
        self.scheduler_config = self._load_scheduler_config()

    def _load_scheduler_config(self) -> Dict:
        """
        Load scheduler configuration from file

        Returns:
            Dictionary with scheduler configuration
        """
        config_path = "../../configs/scheduler.yaml"
        default_config = {
            "email": {
                "enabled": False,
                "smtp_server": "",
                "smtp_port": 587,
                "use_tls": True,
                "username": "",
                "password": "",
                "from_address": "",
                "admin_address": ""
            },
            "schedule": {
                "default_time": "08:00",
                "default_day": "monday",
                "output_dir": "automated_output"
            },
            # Add Excel-specific configuration
            "excel": {
                "visible": False,  # Set to True for debugging
                "cleanup_after_job": True,
                "max_retries": 3
            },
            "jobs": []
        }

        try:
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = yaml.safe_load(f)

                # Ensure Excel configuration exists
                if "excel" not in config:
                    config["excel"] = default_config["excel"]

                logger.info(f"Loaded scheduler configuration from {config_path}")
                return config
            else:
                logger.warning(f"Scheduler configuration not found at {config_path}, using defaults")
                # Create default config file
                os.makedirs(os.path.dirname(config_path), exist_ok=True)
                with open(config_path, 'w', encoding='utf-8') as f:
                    yaml.dump(default_config, f, default_flow_style=False)
                return default_config
        except Exception as e:
            logger.error(f"Error loading scheduler configuration: {e}")
            return default_config
    
    def save_scheduler_config(self) -> bool:
        """
        Save scheduler configuration to file
        
        Returns:
            bool: Success
        """
        config_path = "../../configs/scheduler.yaml"
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                yaml.dump(self.scheduler_config, f, default_flow_style=False)
            logger.info(f"Saved scheduler configuration to {config_path}")
            return True
        except Exception as e:
            logger.error(f"Error saving scheduler configuration: {e}")
            return False
    
    def get_jobs(self) -> List[Dict]:
        """
        Get all scheduled jobs
        
        Returns:
            List of job dictionaries
        """
        return self.scheduler_config.get("jobs", [])
    
    def add_job(self, job_config: Dict) -> bool:
        """
        Add a new scheduled job
        
        Args:
            job_config: Dictionary with job configuration
            
        Returns:
            bool: Success
        """
        try:
            # Validate job config
            required_fields = ["job_id", "analytics_id", "schedule_type", "data_source_pattern"]
            for field in required_fields:
                if field not in job_config:
                    logger.error(f"Missing required field '{field}' in job configuration")
                    return False
            
            # Add job to configuration
            jobs = self.scheduler_config.get("jobs", [])
            
            # Check if job ID already exists
            for i, job in enumerate(jobs):
                if job.get("job_id") == job_config["job_id"]:
                    # Update existing job
                    jobs[i] = job_config
                    self.scheduler_config["jobs"] = jobs
                    self.save_scheduler_config()
                    # Remove old schedule if running
                    if self.running:
                        self.scheduler.clear(tag=job_config["job_id"])
                        self._schedule_job(job_config)
                    return True
            
            # Add new job
            jobs.append(job_config)
            self.scheduler_config["jobs"] = jobs
            self.save_scheduler_config()
            
            # Add to scheduler if running
            if self.running:
                self._schedule_job(job_config)
            
            return True
            
        except Exception as e:
            logger.error(f"Error adding job: {e}")
            return False
    
    def remove_job(self, job_id: str) -> bool:
        """
        Remove a scheduled job
        
        Args:
            job_id: Job identifier
            
        Returns:
            bool: Success
        """
        try:
            jobs = self.scheduler_config.get("jobs", [])
            
            # Find job by ID
            for i, job in enumerate(jobs):
                if job.get("job_id") == job_id:
                    # Remove job
                    del jobs[i]
                    self.scheduler_config["jobs"] = jobs
                    self.save_scheduler_config()
                    
                    # Remove from scheduler if running
                    if self.running:
                        self.scheduler.clear(tag=job_id)
                    
                    return True
            
            logger.warning(f"Job '{job_id}' not found")
            return False
            
        except Exception as e:
            logger.error(f"Error removing job: {e}")
            return False
    
    def update_email_config(self, email_config: Dict) -> bool:
        """
        Update email configuration
        
        Args:
            email_config: Dictionary with email configuration
            
        Returns:
            bool: Success
        """
        try:
            self.scheduler_config["email"] = email_config
            self.save_scheduler_config()
            return True
        except Exception as e:
            logger.error(f"Error updating email configuration: {e}")
            return False
    
    def get_email_config(self) -> Dict:
        """
        Get email configuration
        
        Returns:
            Dictionary with email configuration
        """
        return self.scheduler_config.get("email", {})
    
    def get_schedule_config(self) -> Dict:
        """
        Get schedule configuration
        
        Returns:
            Dictionary with schedule configuration
        """
        return self.scheduler_config.get("schedule", {})
    
    def update_schedule_config(self, schedule_config: Dict) -> bool:
        """
        Update schedule configuration
        
        Args:
            schedule_config: Dictionary with schedule configuration
            
        Returns:
            bool: Success
        """
        try:
            self.scheduler_config["schedule"] = schedule_config
            self.save_scheduler_config()
            return True
        except Exception as e:
            logger.error(f"Error updating schedule configuration: {e}")
            return False
    
    def start_scheduler(self) -> bool:
        """
        Start the scheduler in a background thread
        
        Returns:
            bool: Success
        """
        if self.running:
            logger.warning("Scheduler is already running")
            return False
        
        # Schedule all jobs
        self._schedule_all_jobs()
        
        # Start scheduler in a separate thread
        self.running = True
        self.scheduler_thread = threading.Thread(target=self._run_scheduler, daemon=True)
        self.scheduler_thread.start()
        
        logger.info("Scheduler started")
        return True

    def stop_scheduler(self) -> bool:
        """
        Stop the scheduler

        Returns:
            bool: Success
        """
        if not self.running:
            logger.warning("Scheduler is not running")
            return False

        # Set flag to stop
        self.running = False

        # Clear all scheduled jobs
        self.scheduler.clear()

        # Clean up Excel resources
        self._cleanup_excel_processors()

        # Wait for thread to join
        if self.scheduler_thread and self.scheduler_thread.is_alive():
            self.scheduler_thread.join(timeout=1.0)

        logger.info("Scheduler stopped")
        return True
    
    def is_running(self) -> bool:
        """
        Check if scheduler is running
        
        Returns:
            bool: True if running, False otherwise
        """
        return self.running
    
    def _schedule_all_jobs(self) -> None:
        """Schedule all configured jobs"""
        jobs = self.scheduler_config.get("jobs", [])
        
        for job in jobs:
            self._schedule_job(job)
    
    def _schedule_job(self, job: Dict) -> None:
        """
        Schedule a single job
        
        Args:
            job: Job configuration dictionary
        """
        job_id = job.get("job_id")
        analytics_id = job.get("analytics_id")
        schedule_type = job.get("schedule_type", "daily")
        
        # Get schedule time
        schedule_time = job.get("schedule_time")
        if not schedule_time:
            schedule_time = self.scheduler_config.get("schedule", {}).get("default_time", "08:00")
        
        # Create the job function
        job_func = lambda: self._run_analytics_job(job)
        
        # Schedule based on type
        if schedule_type == "daily":
            self.scheduler.every().day.at(schedule_time).do(job_func).tag(job_id)
            logger.info(f"Scheduled job '{job_id}' for daily execution at {schedule_time}")
            
        elif schedule_type == "weekly":
            day = job.get("schedule_day") or self.scheduler_config.get("schedule", {}).get("default_day", "monday")
            
            if day.lower() == "monday":
                self.scheduler.every().monday.at(schedule_time).do(job_func).tag(job_id)
            elif day.lower() == "tuesday":
                self.scheduler.every().tuesday.at(schedule_time).do(job_func).tag(job_id)
            elif day.lower() == "wednesday":
                self.scheduler.every().wednesday.at(schedule_time).do(job_func).tag(job_id)
            elif day.lower() == "thursday":
                self.scheduler.every().thursday.at(schedule_time).do(job_func).tag(job_id)
            elif day.lower() == "friday":
                self.scheduler.every().friday.at(schedule_time).do(job_func).tag(job_id)
            elif day.lower() == "saturday":
                self.scheduler.every().saturday.at(schedule_time).do(job_func).tag(job_id)
            elif day.lower() == "sunday":
                self.scheduler.every().sunday.at(schedule_time).do(job_func).tag(job_id)
            
            logger.info(f"Scheduled job '{job_id}' for weekly execution on {day} at {schedule_time}")
            
        elif schedule_type == "monthly":
            day_of_month = job.get("schedule_day", 1)
            
            # Schedule for a specific day of month
            def monthly_job():
                # Only run on the specified day of the month
                if datetime.datetime.now().day == int(day_of_month):
                    job_func()
            
            # Run check daily
            self.scheduler.every().day.at(schedule_time).do(monthly_job).tag(job_id)
            
            logger.info(f"Scheduled job '{job_id}' for monthly execution on day {day_of_month} at {schedule_time}")
    
    def _run_scheduler(self) -> None:
        """Run the scheduler loop"""
        while self.running:
            try:
                self.scheduler.run_pending()
                time.sleep(1)
            except Exception as e:
                logger.error(f"Error in scheduler loop: {e}")
                time.sleep(5)  # Sleep a bit longer after error

    def _run_analytics_job(self, job: Dict) -> None:
        """
        Run an analytics job

        Args:
            job: Job configuration dictionary
        """
        job_id = job.get("job_id")
        analytics_id = job.get("analytics_id")
        data_source_pattern = job.get("data_source_pattern")
        recipients = job.get("email_recipients", [])

        logger.info(f"Running scheduled job '{job_id}' for analytics ID {analytics_id}")

        excel_processor = None

        try:
            # Get configuration
            config = self.config_manager.get_config(analytics_id)
            if not config:
                logger.error(f"Configuration for QA-{analytics_id} not found")
                return

            # Find matching data files
            data_files = self._find_data_files(data_source_pattern)
            if not data_files:
                logger.warning(f"No data files found matching pattern '{data_source_pattern}'")
                self._send_error_notification(
                    job_id=job_id,
                    analytics_id=analytics_id,
                    error_message=f"No data files found matching pattern '{data_source_pattern}'",
                    recipients=recipients
                )
                return

            # Use the most recent file
            data_file = data_files[0]
            logger.info(f"Using most recent data file: {data_file}")

            # Create output directory
            output_dir = self.scheduler_config.get("schedule", {}).get("output_dir", "automated_output")
            job_output_dir = os.path.join(output_dir, f"job_{job_id}")
            os.makedirs(job_output_dir, exist_ok=True)

            # Check if any validation uses Excel formulas
            uses_excel_formulas = self._check_for_excel_formulas(config)

            if uses_excel_formulas:
                # Initialize Excel processor
                excel_config = self.scheduler_config.get("excel", {})
                visible = excel_config.get("visible", False)

                logger.info(f"Initializing Excel processor for job '{job_id}'")
                excel_processor = ExcelFormulaProcessor(visible=visible)

                # Store processor for tracking
                self.excel_processors[job_id] = excel_processor

            # Process data
            processor = self.data_processor_class(config)

            # If using Excel formulas, pass the processor to the data processor
            if uses_excel_formulas and excel_processor:
                # Assuming the data processor has a way to accept an Excel processor
                # This would need to be implemented in the data processor
                processor.excel_processor = excel_processor

            success, message = processor.process_data(data_file)

            if not success:
                logger.error(f"Processing failed for job '{job_id}': {message}")
                self._send_error_notification(
                    job_id=job_id,
                    analytics_id=analytics_id,
                    error_message=message,
                    recipients=recipients
                )
                return

            # Generate reports
            report_generator = self.report_generator_class(config, processor.results)

            # Generate main report
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            main_report_path = os.path.join(
                job_output_dir,
                f"QA_{analytics_id}_Main_{timestamp}.xlsx"
            )

            main_report = report_generator.generate_main_report(main_report_path, data_file)

            if not main_report:
                logger.error(f"Failed to generate main report for job '{job_id}'")
                self._send_error_notification(
                    job_id=job_id,
                    analytics_id=analytics_id,
                    error_message="Failed to generate main report",
                    recipients=recipients
                )
                return

            # Generate individual reports if configured
            individual_reports = []
            if job.get("generate_individual_reports", False):
                individual_reports = report_generator.generate_individual_reports(output_dir=job_output_dir)

            # Send email notification if configured
            if job.get("send_email", False) and recipients:
                self._send_email_notification(
                    job_id=job_id,
                    analytics_id=analytics_id,
                    config=config,
                    main_report_path=main_report,
                    individual_reports=individual_reports,
                    recipients=recipients,
                    results=processor.results
                )

            logger.info(f"Job '{job_id}' completed successfully")

        except Exception as e:
            logger.error(f"Error running job '{job_id}': {e}")
            self._send_error_notification(
                job_id=job_id,
                analytics_id=analytics_id,
                error_message=str(e),
                recipients=recipients
            )
        finally:
            # Clean up Excel resources
            if excel_processor:
                try:
                    excel_config = self.scheduler_config.get("excel", {})
                    if excel_config.get("cleanup_after_job", True):
                        logger.info(f"Cleaning up Excel resources for job '{job_id}'")
                        excel_processor.cleanup()

                    # Remove from tracking
                    if job_id in self.excel_processors:
                        del self.excel_processors[job_id]
                except Exception as e:
                    logger.warning(f"Error cleaning up Excel resources for job '{job_id}': {e}")

    def _check_for_excel_formulas(self, config: Dict) -> bool:
        """
        Check if the configuration contains any Excel formula validations

        Args:
            config: Analytics configuration

        Returns:
            bool: True if the configuration uses Excel formulas
        """
        if "validations" not in config:
            return False

        # Check each validation
        for validation in config.get("validations", []):
            # Check for custom_formula rule type
            if validation.get("rule") == "custom_formula":
                return True

            # Check for original_formula in parameters
            parameters = validation.get("parameters", {})
            if "original_formula" in parameters:
                return True

        return False

    def _cleanup_excel_processors(self) -> None:
        """Clean up all Excel processors"""
        logger.info("Cleaning up all Excel processors")

        for job_id, processor in list(self.excel_processors.items()):
            try:
                processor.cleanup()
                logger.info(f"Cleaned up Excel processor for job '{job_id}'")
            except Exception as e:
                logger.warning(f"Error cleaning up Excel processor for job '{job_id}': {e}")

            # Remove from tracking
            del self.excel_processors[job_id]

        # Final check to ensure Excel processes are closed
        try:
            ensure_excel_closed()
        except Exception as e:
            logger.warning(f"Error ensuring Excel processes are closed: {e}")

    def _find_data_files(self, pattern: str) -> List[str]:
        """
        Find data files matching the pattern
        
        Args:
            pattern: File pattern to match
            
        Returns:
            List of matching file paths, sorted by modification time (newest first)
        """
        # Handle directory in pattern
        directory = os.path.dirname(pattern)
        if not directory:
            directory = "."
        
        filename_pattern = os.path.basename(pattern)
        
        # Ensure directory exists
        if not os.path.exists(directory):
            logger.warning(f"Directory not found: {directory}")
            return []
        
        # Find all files in directory
        all_files = []
        for file in os.listdir(directory):
            if fnmatch.fnmatch(file, filename_pattern):
                file_path = os.path.join(directory, file)
                if os.path.isfile(file_path):
                    all_files.append(file_path)
        
        # Sort by modification time (newest first)
        all_files.sort(key=os.path.getmtime, reverse=True)
        
        return all_files
    
    def _send_email_notification(
        self, job_id: str, analytics_id: str, config: Dict, 
        main_report_path: str, individual_reports: List[str], 
        recipients: List[str], results: Dict
    ) -> None:
        """
        Send email notification with report
        
        Args:
            job_id: Job identifier
            analytics_id: Analytics ID
            config: Analytics configuration
            main_report_path: Path to main report
            individual_reports: List of paths to individual reports
            recipients: List of email recipients
            results: Job results dictionary
        """
        email_config = self.scheduler_config.get("email", {})
        
        if not email_config.get("enabled", False):
            logger.warning("Email notifications are disabled")
            return
        
        try:
            # Create message
            msg = MIMEMultipart()
            msg['From'] = email_config.get("from_address")
            msg['To'] = ", ".join(recipients)
            msg['Subject'] = f"QA Analytics Report: {config.get('analytic_name', f'QA-{analytics_id}')}"
            
            # Add email body
            body = self._generate_email_body(job_id, analytics_id, config, results)
            msg.attach(MIMEText(body, 'html'))
            
            # Attach main report
            if os.path.exists(main_report_path):
                with open(main_report_path, 'rb') as f:
                    attachment = MIMEApplication(f.read(), _subtype="xlsx")
                    attachment.add_header('Content-Disposition', 'attachment', 
                                       filename=os.path.basename(main_report_path))
                    msg.attach(attachment)
            
            # Add individual reports if needed (limit to 5 to avoid large emails)
            if individual_reports and len(individual_reports) <= 5:
                for report_path in individual_reports:
                    if os.path.exists(report_path):
                        with open(report_path, 'rb') as f:
                            attachment = MIMEApplication(f.read(), _subtype="xlsx")
                            attachment.add_header('Content-Disposition', 'attachment', 
                                               filename=os.path.basename(report_path))
                            msg.attach(attachment)
            
            # Send email
            server = smtplib.SMTP(email_config.get("smtp_server"), email_config.get("smtp_port", 587))
            
            if email_config.get("use_tls", True):
                server.starttls()
            
            if email_config.get("username") and email_config.get("password"):
                server.login(email_config.get("username"), email_config.get("password"))
            
            server.send_message(msg)
            server.quit()
            
            logger.info(f"Email notification sent for job '{job_id}'")
            
        except Exception as e:
            logger.error(f"Error sending email notification for job '{job_id}': {e}")
    
    def _send_error_notification(
        self, job_id: str, analytics_id: str, 
        error_message: str, recipients: List[str]
    ) -> None:
        """
        Send error notification email
        
        Args:
            job_id: Job identifier
            analytics_id: Analytics ID
            error_message: Error message
            recipients: List of email recipients
        """
        email_config = self.scheduler_config.get("email", {})
        
        if not email_config.get("enabled", False):
            logger.warning("Email notifications are disabled")
            return
        
        try:
            # Create message
            msg = MIMEMultipart()
            msg['From'] = email_config.get("from_address")
            msg['To'] = ", ".join(recipients)
            msg['Subject'] = f"ERROR: QA Analytics Job {job_id} (QA-{analytics_id})"
            
            # Create HTML body
            body = f"""
            <html>
            <head>
                <style>
                    body {{ font-family: Arial, sans-serif; }}
                    .header {{ background-color: #cc0000; color: white; padding: 10px; }}
                    .content {{ margin: 20px 0; }}
                    .footer {{ margin-top: 20px; font-size: 0.8em; color: #666; }}
                </style>
            </head>
            <body>
                <div class="header">
                    <h2>QA Analytics Error Notification</h2>
                </div>
                
                <div class="content">
                    <h3>Error in Job {job_id} (QA-{analytics_id})</h3>
                    <p>The following error occurred during job execution:</p>
                    <p style="background-color: #f8f8f8; padding: 10px; border-left: 4px solid #cc0000;">
                        {error_message}
                    </p>
                    <p>Please check the job configuration and data source.</p>
                    <p>Time: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
                </div>
                
                <div class="footer">
                    <p>This is an automated notification from the QA Analytics system. Please do not reply to this email.</p>
                </div>
            </body>
            </html>
            """
            
            msg.attach(MIMEText(body, 'html'))
            
            # Send email
            server = smtplib.SMTP(email_config.get("smtp_server"), email_config.get("smtp_port", 587))
            
            if email_config.get("use_tls", True):
                server.starttls()
            
            if email_config.get("username") and email_config.get("password"):
                server.login(email_config.get("username"), email_config.get("password"))
            
            server.send_message(msg)
            server.quit()
            
            logger.info(f"Error notification sent for job '{job_id}'")
            
        except Exception as e:
            logger.error(f"Error sending error notification for job '{job_id}': {e}")
    
    def _generate_email_body(self, job_id: str, analytics_id: str, config: Dict, results: Dict) -> str:
        """
        Generate HTML email body
        
        Args:
            job_id: Job identifier
            analytics_id: Analytics ID
            config: Analytics configuration
            results: Job results dictionary
            
        Returns:
            HTML email body
        """
        # Get summary data
        summary = results.get('summary')
        detail = results.get('detail')
        warnings = results.get('warnings', [])
        
        # Calculate overall statistics
        total_records = len(detail) if detail is not None else 0
        gc_count = sum(detail['Compliance'] == 'GC') if detail is not None and 'Compliance' in detail else 0
        dnc_count = sum(detail['Compliance'] == 'DNC') if detail is not None and 'Compliance' in detail else 0
        pc_count = sum(detail['Compliance'] == 'PC') if detail is not None and 'Compliance' in detail else 0
        
        error_pct = (dnc_count / total_records * 100) if total_records > 0 else 0
        threshold = config.get('thresholds', {}).get('error_percentage', 5.0)
        threshold_status = "EXCEEDS THRESHOLD" if error_pct > threshold else "Within Threshold"
        
        # Format timestamp
        timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # Create HTML body
        html = f"""
        <html>
        <head>
            <style>
                body {{ font-family: Arial, sans-serif; }}
                .header {{ background-color: #003366; color: white; padding: 10px; }}
                .summary {{ margin: 20px 0; }}
                table {{ border-collapse: collapse; width: 100%; }}
                th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
                th {{ background-color: #f2f2f2; }}
                .exceeds {{ background-color: #ffcccc; }}
                .within {{ background-color: #ccffcc; }}
                .warning {{ background-color: #fff8e1; }}
                .footer {{ margin-top: 20px; font-size: 0.8em; color: #666; }}
            </style>
        </head>
        <body>
            <div class="header">
                <h2>QA Analytics Automated Report</h2>
            </div>
            
            <div class="summary">
                <h3>Summary for {config.get('analytic_name', f'QA-{analytics_id}')}</h3>
                <p>This report was generated automatically by the QA Analytics Automation Scheduler.</p>
                
                <table>
                    <tr>
                        <th>Job ID</th>
                        <td>{job_id}</td>
                    </tr>
                    <tr>
                        <th>Analytics ID</th>
                        <td>QA-{analytics_id}</td>
                    </tr>
                    <tr>
                        <th>Run Date</th>
                        <td>{timestamp}</td>
                    </tr>
                    <tr>
                        <th>Total Records</th>
                        <td>{total_records}</td>
                    </tr>
                        <tr>
                            <th>Generally Conforms (GC)</th>
                            <td>{f"{gc_count} ({(gc_count/total_records*100):.1f}%)" if total_records > 0 else f"{gc_count} (0%)"}</td>
                        </tr>
                        <tr>
                            <th>Does Not Conform (DNC)</th>
                            <td>{f"{dnc_count} ({(dnc_count/total_records*100):.1f}%)" if total_records > 0 else f"{dnc_count} (0%)"}</td>
                        </tr>
                        <tr>
                            <th>Partially Conforms (PC)</th>
                            <td>{f"{pc_count} ({(pc_count/total_records*100):.1f}%)" if total_records > 0 else f"{pc_count} (0%)"}</td>
                        </tr>

                    <tr>
                        <th>Error Threshold</th>
                        <td>{threshold}%</td>
                    </tr>
                    <tr class="{'exceeds' if error_pct > threshold else 'within'}">
                        <th>Threshold Status</th>
                        <td>{threshold_status}</td>
                    </tr>
                </table>
            </div>
        """
        
        # Add warnings if any
        if warnings:
            html += """
            <div class="warnings">
                <h3>Warnings</h3>
                <table class="warning">
                    <tr>
                        <th>Warning</th>
                    </tr>
            """
            
            for warning in warnings:
                html += f"""
                    <tr>
                        <td>{warning}</td>
                    </tr>
                """
            
            html += """
                </table>
            </div>
            """
        
        # Add group summary if available
        if summary is not None and not summary.empty:
            html += """
            <div class="group-summary">
                <h3>Group Summary</h3>
                <table>
                    <tr>
                        <th>Group</th>
                        <th>GC</th>
                        <th>PC</th>
                        <th>DNC</th>
                        <th>Total</th>
                        <th>DNC %</th>
                        <th>Status</th>
                    </tr>
            """
            
            # Get group by field
            group_field = config.get('reporting', {}).get('group_by', '')
            
            for _, row in summary.iterrows():
                group_value = row[group_field] if group_field in row else "Unknown"
                gc = row.get('GC', 0)
                pc = row.get('PC', 0)
                dnc = row.get('DNC', 0)
                total = row.get('Total', 0)
                dnc_pct = row.get('DNC_Percentage', 0)
                exceeds = row.get('Exceeds_Threshold', False)
                
                html += f"""
                    <tr class="{'exceeds' if exceeds else 'within'}">
                        <td>{group_value}</td>
                        <td>{gc}</td>
                        <td>{pc}</td>
                        <td>{dnc}</td>
                        <td>{total}</td>
                        <td>{dnc_pct:.2f}%</td>
                        <td>{"EXCEEDS THRESHOLD" if exceeds else "Within Threshold"}</td>
                    </tr>
                """
            
            html += """
                </table>
            </div>
            """
        
        # Add footer
        html += """
            <div class="footer">
                <p>This is an automated notification from the QA Analytics system. Please do not reply to this email.</p>
                <p>The complete report is attached to this email as an Excel file.</p>
                <p>For more details, open the attached report(s) or log into the QA Analytics system.</p>
            </div>
        </body>
        </html>
        """
        
        return html
    
    def run_job_now(self, job_id: str) -> bool:
        """
        Run a job immediately
        
        Args:
            job_id: Job identifier
            
        Returns:
            bool: Success
        """
        jobs = self.scheduler_config.get("jobs", [])
        
        # Find job by ID
        job = None
        for j in jobs:
            if j.get("job_id") == job_id:
                job = j
                break
        
        if not job:
            logger.warning(f"Job '{job_id}' not found")
            return False
        
        # Run job in a separate thread
        threading.Thread(target=self._run_analytics_job, args=(job,), daemon=True).start()
        logger.info(f"Started job '{job_id}' manually")
        return True

    def test_email_configuration(self, recipient: str) -> Tuple[bool, str]:
        """
        Test email configuration by sending a test email
        
        Args:
            recipient: Recipient email address
            
        Returns:
            Tuple of (success, message)
        """
        email_config = self.scheduler_config.get("email", {})
        
        if not email_config.get("enabled", False):
            return False, "Email notifications are disabled"
        
        if not email_config.get("smtp_server"):
            return False, "SMTP server not configured"
        
        if not email_config.get("from_address"):
            return False, "From address not configured"
        
        try:
            # Create message
            msg = MIMEMultipart()
            msg['From'] = email_config.get("from_address")
            msg['To'] = recipient
            msg['Subject'] = "QA Analytics Test Email"
            
            # Create HTML body
            body = """
            <html>
            <head>
                <style>
                    body { font-family: Arial, sans-serif; }
                    .header { background-color: #003366; color: white; padding: 10px; }
                    .content { margin: 20px 0; }
                    .footer { margin-top: 20px; font-size: 0.8em; color: #666; }
                </style>
            </head>
            <body>
                <div class="header">
                    <h2>QA Analytics Test Email</h2>
                </div>
                
                <div class="content">
                    <p>This is a test email from the QA Analytics Automation Scheduler.</p>
                    <p>If you received this email, your email configuration is working correctly.</p>
                </div>
                
                <div class="footer">
                    <p>This is an automated test from the QA Analytics system. Please do not reply to this email.</p>
                </div>
            </body>
            </html>
            """
            
            msg.attach(MIMEText(body, 'html'))
            
            # Send email
            server = smtplib.SMTP(email_config.get("smtp_server"), email_config.get("smtp_port", 587))
            
            if email_config.get("use_tls", True):
                server.starttls()
            
            if email_config.get("username") and email_config.get("password"):
                server.login(email_config.get("username"), email_config.get("password"))
            
            server.send_message(msg)
            server.quit()
            
            logger.info(f"Test email sent to {recipient}")
            return True, f"Test email sent to {recipient}"
            
        except Exception as e:
            logger.error(f"Error sending test email: {e}")
            return False, f"Error sending test email: {e}"


# Add the SchedulerUI class to automation_scheduler.py
# This should be appended at the end of the file

class SchedulerUI:
    """User interface for managing scheduled jobs"""

    def __init__(self, parent_frame, scheduler):
        """
        Initialize the scheduler UI

        Args:
            parent_frame: Parent tkinter frame
            scheduler: AutomationScheduler instance
        """
        import tkinter as tk
        from tkinter import ttk, messagebox, filedialog

        self.parent = parent_frame
        self.scheduler = scheduler
        self.tk = tk
        self.ttk = ttk
        self.messagebox = messagebox
        self.filedialog = filedialog

        # State variables
        self.job_id_var = tk.StringVar()
        self.analytics_id_var = tk.StringVar()
        self.schedule_type_var = tk.StringVar(value="daily")
        self.schedule_time_var = tk.StringVar(value="08:00")
        self.schedule_day_var = tk.StringVar(value="monday")
        self.data_source_var = tk.StringVar()
        self.send_email_var = tk.BooleanVar(value=False)
        self.generate_individual_var = tk.BooleanVar(value=False)

        # Email configuration variables
        self.email_enabled_var = tk.BooleanVar(value=False)
        self.smtp_server_var = tk.StringVar()
        self.smtp_port_var = tk.StringVar(value="587")
        self.use_tls_var = tk.BooleanVar(value=True)
        self.username_var = tk.StringVar()
        self.password_var = tk.StringVar()
        self.from_address_var = tk.StringVar()
        self.admin_address_var = tk.StringVar()

        # Schedule configuration variables
        self.default_time_var = tk.StringVar(value="08:00")
        self.default_day_var = tk.StringVar(value="monday")
        self.output_dir_var = tk.StringVar(value="automated_output")

        # Email recipients for job
        self.recipients = []

        # Set up the UI
        self._setup_ui()

        # Load configuration
        self._load_config()

    def _setup_ui(self):
        """Set up the scheduler UI"""
        main_frame = self.ttk.Frame(self.parent)
        main_frame.pack(fill=self.tk.BOTH, expand=True, padx=10, pady=10)

        # Create notebook for tabs
        notebook = self.ttk.Notebook(main_frame)
        notebook.pack(fill=self.tk.BOTH, expand=True)

        # Create tabs
        jobs_tab = self.ttk.Frame(notebook)
        config_tab = self.ttk.Frame(notebook)

        notebook.add(jobs_tab, text="Scheduled Jobs")
        notebook.add(config_tab, text="Configuration")

        # Set up jobs tab
        self._setup_jobs_tab(jobs_tab)

        # Set up configuration tab
        self._setup_config_tab(config_tab)

    def _setup_jobs_tab(self, parent):
        """Set up the jobs tab"""
        # Split into left (job list) and right (job details) panes
        paned = self.ttk.PanedWindow(parent, orient=self.tk.HORIZONTAL)
        paned.pack(fill=self.tk.BOTH, expand=True, padx=5, pady=5)

        # Left pane - job list
        left_frame = self.ttk.Frame(paned)
        paned.add(left_frame, weight=1)

        # Job list frame
        list_frame = self.ttk.LabelFrame(left_frame, text="Scheduled Jobs")
        list_frame.pack(fill=self.tk.BOTH, expand=True, padx=5, pady=5)

        # Job list treeview
        columns = ("ID", "Analytics", "Schedule", "Email")
        self.job_tree = self.ttk.Treeview(list_frame, columns=columns, show="headings", height=10)

        # Configure columns
        self.job_tree.column("ID", width=70)
        self.job_tree.column("Analytics", width=100)
        self.job_tree.column("Schedule", width=150)
        self.job_tree.column("Email", width=70, anchor=self.tk.CENTER)

        # Configure headings
        self.job_tree.heading("ID", text="Job ID")
        self.job_tree.heading("Analytics", text="Analytics ID")
        self.job_tree.heading("Schedule", text="Schedule")
        self.job_tree.heading("Email", text="Email")

        # Add scrollbar
        tree_scroll = self.ttk.Scrollbar(list_frame, orient="vertical", command=self.job_tree.yview)
        self.job_tree.configure(yscrollcommand=tree_scroll.set)

        # Pack tree and scrollbar
        self.job_tree.pack(side=self.tk.LEFT, fill=self.tk.BOTH, expand=True)
        tree_scroll.pack(side=self.tk.RIGHT, fill=self.tk.Y)

        # Bind selection event
        self.job_tree.bind("<<TreeviewSelect>>", self._on_job_selected)

        # Action buttons
        action_frame = self.ttk.Frame(left_frame)
        action_frame.pack(fill=self.tk.X, padx=5, pady=5)

        self.ttk.Button(action_frame, text="New Job", command=self._new_job).pack(side=self.tk.LEFT)
        self.ttk.Button(action_frame, text="Delete Job", command=self._delete_job).pack(side=self.tk.LEFT, padx=5)
        self.ttk.Button(action_frame, text="Run Now", command=self._run_job_now).pack(side=self.tk.LEFT)

        # Right pane - job details
        right_frame = self.ttk.Frame(paned)
        paned.add(right_frame, weight=2)

        # Job details frame
        details_frame = self.ttk.LabelFrame(right_frame, text="Job Details")
        details_frame.pack(fill=self.tk.BOTH, expand=True, padx=5, pady=5)

        # Create scrollable frame for details
        canvas = self.tk.Canvas(details_frame)
        scrollbar = self.ttk.Scrollbar(details_frame, orient="vertical", command=canvas.yview)

        scroll_frame = self.ttk.Frame(canvas)
        scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=self.tk.LEFT, fill=self.tk.BOTH, expand=True)
        scrollbar.pack(side=self.tk.RIGHT, fill=self.tk.Y)

        # Job ID
        self.ttk.Label(scroll_frame, text="Job ID:").grid(row=0, column=0, sticky=self.tk.W, padx=5, pady=5)
        self.ttk.Entry(scroll_frame, textvariable=self.job_id_var, width=20).grid(row=0, column=1, sticky=self.tk.W,
                                                                                  padx=5, pady=5)

        # Analytics ID
        self.ttk.Label(scroll_frame, text="Analytics ID:").grid(row=1, column=0, sticky=self.tk.W, padx=5, pady=5)

        # Get available analytics
        analytics_frame = self.ttk.Frame(scroll_frame)
        analytics_frame.grid(row=1, column=1, sticky=self.tk.W, padx=5, pady=5)

        analytics_entry = self.ttk.Entry(analytics_frame, textvariable=self.analytics_id_var, width=20)
        analytics_entry.pack(side=self.tk.LEFT)

        # Or use combobox if ConfigManager is available
        try:
            analytics = self.scheduler.config_manager.get_available_analytics()
            if analytics:
                analytics_values = [id for id, _ in analytics]
                self.analytics_id_var.set(analytics_values[0] if analytics_values else "")

                analytics_entry.destroy()
                analytics_combo = self.ttk.Combobox(analytics_frame, textvariable=self.analytics_id_var,
                                                    values=analytics_values, state="readonly", width=20)
                analytics_combo.pack(side=self.tk.LEFT)
        except:
            pass

        # Schedule type
        self.ttk.Label(scroll_frame, text="Schedule Type:").grid(row=2, column=0, sticky=self.tk.W, padx=5, pady=5)

        schedule_frame = self.ttk.Frame(scroll_frame)
        schedule_frame.grid(row=2, column=1, sticky=self.tk.W, padx=5, pady=5)

        self.ttk.Radiobutton(schedule_frame, text="Daily", variable=self.schedule_type_var,
                             value="daily", command=self._update_schedule_options).pack(side=self.tk.LEFT)
        self.ttk.Radiobutton(schedule_frame, text="Weekly", variable=self.schedule_type_var,
                             value="weekly", command=self._update_schedule_options).pack(side=self.tk.LEFT,
                                                                                         padx=(10, 0))
        self.ttk.Radiobutton(schedule_frame, text="Monthly", variable=self.schedule_type_var,
                             value="monthly", command=self._update_schedule_options).pack(side=self.tk.LEFT,
                                                                                          padx=(10, 0))

        # Schedule time
        self.ttk.Label(scroll_frame, text="Time (HH:MM):").grid(row=3, column=0, sticky=self.tk.W, padx=5, pady=5)
        self.ttk.Entry(scroll_frame, textvariable=self.schedule_time_var, width=10).grid(row=3, column=1,
                                                                                         sticky=self.tk.W, padx=5,
                                                                                         pady=5)

        # Day options (for weekly and monthly)
        self.day_label = self.ttk.Label(scroll_frame, text="Day:")
        self.day_label.grid(row=4, column=0, sticky=self.tk.W, padx=5, pady=5)

        self.day_frame = self.ttk.Frame(scroll_frame)
        self.day_frame.grid(row=4, column=1, sticky=self.tk.W, padx=5, pady=5)

        # Weekly options
        self.week_frame = self.ttk.Frame(self.day_frame)

        days = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
        self.day_combo = self.ttk.Combobox(self.week_frame, textvariable=self.schedule_day_var,
                                           values=days, state="readonly", width=15)
        self.day_combo.pack(side=self.tk.LEFT)

        # Monthly options
        self.month_frame = self.ttk.Frame(self.day_frame)

        days_of_month = [str(i) for i in range(1, 29)]  # 1-28
        self.day_of_month_combo = self.ttk.Combobox(self.month_frame, textvariable=self.schedule_day_var,
                                                    values=days_of_month, state="readonly", width=5)
        self.day_of_month_combo.pack(side=self.tk.LEFT)
        self.ttk.Label(self.month_frame, text="of each month").pack(side=self.tk.LEFT, padx=(5, 0))

        # Show appropriate day options based on schedule type
        self._update_schedule_options()

        # Data source pattern
        self.ttk.Label(scroll_frame, text="Data File Pattern:").grid(row=5, column=0, sticky=self.tk.W, padx=5, pady=5)

        data_frame = self.ttk.Frame(scroll_frame)
        data_frame.grid(row=5, column=1, sticky=self.tk.W, padx=5, pady=5)

        self.ttk.Entry(data_frame, textvariable=self.data_source_var, width=40).pack(side=self.tk.LEFT)
        self.ttk.Button(data_frame, text="Browse...", command=self._browse_data_pattern).pack(side=self.tk.LEFT,
                                                                                              padx=(5, 0))

        # Email options
        self.ttk.Checkbutton(scroll_frame, text="Send Email Notification", variable=self.send_email_var,
                             command=self._update_email_options).grid(row=6, column=0, columnspan=2, sticky=self.tk.W,
                                                                      padx=5, pady=5)

        # Recipients frame
        self.recipients_frame = self.ttk.LabelFrame(scroll_frame, text="Email Recipients")
        self.recipients_frame.grid(row=7, column=0, columnspan=2, sticky=self.tk.EW, padx=5, pady=5)

        # Recipients list
        self.recipients_list = self.tk.Listbox(self.recipients_frame, height=4, width=40)
        recipients_scroll = self.ttk.Scrollbar(self.recipients_frame, orient="vertical",
                                               command=self.recipients_list.yview)
        self.recipients_list.configure(yscrollcommand=recipients_scroll.set)

        self.recipients_list.pack(side=self.tk.LEFT, fill=self.tk.BOTH, expand=True, padx=5, pady=5)
        recipients_scroll.pack(side=self.tk.RIGHT, fill=self.tk.Y, pady=5)

        # Recipients actions
        recipients_actions = self.ttk.Frame(self.recipients_frame)
        recipients_actions.pack(fill=self.tk.X, padx=5, pady=5)

        self.recipient_var = self.tk.StringVar()
        self.ttk.Entry(recipients_actions, textvariable=self.recipient_var, width=30).pack(side=self.tk.LEFT)
        self.ttk.Button(recipients_actions, text="Add", command=self._add_recipient).pack(side=self.tk.LEFT,
                                                                                          padx=(5, 0))
        self.ttk.Button(recipients_actions, text="Remove", command=self._remove_recipient).pack(side=self.tk.LEFT,
                                                                                                padx=(5, 0))

        # Individual reports option
        self.ttk.Checkbutton(scroll_frame, text="Generate Individual Reports",
                             variable=self.generate_individual_var).grid(row=8, column=0, columnspan=2,
                                                                         sticky=self.tk.W, padx=5, pady=5)

        # Save button
        save_frame = self.ttk.Frame(scroll_frame)
        save_frame.grid(row=9, column=0, columnspan=2, sticky=self.tk.E, padx=5, pady=(15, 5))

        self.ttk.Button(save_frame, text="Save Job", command=self._save_job).pack(side=self.tk.RIGHT)

        # Initially hide recipients frame
        self.recipients_frame.grid_remove()

        # Scheduler status frame
        status_frame = self.ttk.LabelFrame(right_frame, text="Scheduler Status")
        status_frame.pack(fill=self.tk.X, padx=5, pady=5)

        status_inner = self.ttk.Frame(status_frame)
        status_inner.pack(fill=self.tk.X, padx=10, pady=10)

        self.status_label = self.ttk.Label(status_inner, text="Scheduler is not running")
        self.status_label.pack(side=self.tk.LEFT)

        self.start_btn = self.ttk.Button(status_inner, text="Start Scheduler", command=self._start_scheduler)
        self.start_btn.pack(side=self.tk.RIGHT)

        self.stop_btn = self.ttk.Button(status_inner, text="Stop Scheduler", command=self._stop_scheduler)
        self.stop_btn.pack(side=self.tk.RIGHT, padx=(0, 5))
        self.stop_btn.config(state=self.tk.DISABLED)

        # Update scheduler status display
        self._update_scheduler_status()

    def _setup_config_tab(self, parent):
        """Set up the configuration tab"""
        # Create notebook for configuration tabs
        notebook = self.ttk.Notebook(parent)
        notebook.pack(fill=self.tk.BOTH, expand=True, padx=5, pady=5)

        # Create tabs
        email_tab = self.ttk.Frame(notebook)
        schedule_tab = self.ttk.Frame(notebook)

        notebook.add(email_tab, text="Email Settings")
        notebook.add(schedule_tab, text="Schedule Settings")

        # Email configuration
        email_frame = self.ttk.LabelFrame(email_tab, text="Email Configuration")
        email_frame.pack(fill=self.tk.BOTH, expand=True, padx=10, pady=10)

        # Enable email
        self.ttk.Checkbutton(email_frame, text="Enable Email Notifications",
                             variable=self.email_enabled_var,
                             command=self._update_email_config_state).grid(row=0, column=0, columnspan=2,
                                                                           sticky=self.tk.W, padx=5, pady=5)

        # SMTP Server
        self.ttk.Label(email_frame, text="SMTP Server:").grid(row=1, column=0, sticky=self.tk.W, padx=5, pady=5)
        self.smtp_server_entry = self.ttk.Entry(email_frame, textvariable=self.smtp_server_var, width=30)
        self.smtp_server_entry.grid(row=1, column=1, sticky=self.tk.W, padx=5, pady=5)

        # SMTP Port
        self.ttk.Label(email_frame, text="SMTP Port:").grid(row=2, column=0, sticky=self.tk.W, padx=5, pady=5)
        self.smtp_port_entry = self.ttk.Entry(email_frame, textvariable=self.smtp_port_var, width=10)
        self.smtp_port_entry.grid(row=2, column=1, sticky=self.tk.W, padx=5, pady=5)

        # Use TLS
        self.use_tls_check = self.ttk.Checkbutton(email_frame, text="Use TLS", variable=self.use_tls_var)
        self.use_tls_check.grid(row=3, column=0, columnspan=2, sticky=self.tk.W, padx=5, pady=5)

        # Username
        self.ttk.Label(email_frame, text="Username:").grid(row=4, column=0, sticky=self.tk.W, padx=5, pady=5)
        self.username_entry = self.ttk.Entry(email_frame, textvariable=self.username_var, width=30)
        self.username_entry.grid(row=4, column=1, sticky=self.tk.W, padx=5, pady=5)

        # Password
        self.ttk.Label(email_frame, text="Password:").grid(row=5, column=0, sticky=self.tk.W, padx=5, pady=5)
        self.password_entry = self.ttk.Entry(email_frame, textvariable=self.password_var, width=30, show="*")
        self.password_entry.grid(row=5, column=1, sticky=self.tk.W, padx=5, pady=5)

        # From Address
        self.ttk.Label(email_frame, text="From Address:").grid(row=6, column=0, sticky=self.tk.W, padx=5, pady=5)
        self.from_address_entry = self.ttk.Entry(email_frame, textvariable=self.from_address_var, width=30)
        self.from_address_entry.grid(row=6, column=1, sticky=self.tk.W, padx=5, pady=5)

        # Admin Address
        self.ttk.Label(email_frame, text="Admin Address:").grid(row=7, column=0, sticky=self.tk.W, padx=5, pady=5)
        self.admin_address_entry = self.ttk.Entry(email_frame, textvariable=self.admin_address_var, width=30)
        self.admin_address_entry.grid(row=7, column=1, sticky=self.tk.W, padx=5, pady=5)

        # Test email button
        test_frame = self.ttk.Frame(email_frame)
        test_frame.grid(row=8, column=0, columnspan=2, sticky=self.tk.E, padx=5, pady=(15, 5))

        self.test_email_btn = self.ttk.Button(test_frame, text="Test Email", command=self._test_email)
        self.test_email_btn.pack(side=self.tk.LEFT)

        self.save_email_btn = self.ttk.Button(test_frame, text="Save Email Config", command=self._save_email_config)
        self.save_email_btn.pack(side=self.tk.RIGHT, padx=(10, 0))

        # Schedule configuration
        schedule_frame = self.ttk.LabelFrame(schedule_tab, text="Schedule Configuration")
        schedule_frame.pack(fill=self.tk.BOTH, expand=True, padx=10, pady=10)

        # Default time
        self.ttk.Label(schedule_frame, text="Default Time (HH:MM):").grid(row=0, column=0, sticky=self.tk.W, padx=5,
                                                                          pady=5)
        self.ttk.Entry(schedule_frame, textvariable=self.default_time_var, width=10).grid(row=0, column=1,
                                                                                          sticky=self.tk.W, padx=5,
                                                                                          pady=5)

        # Default day
        self.ttk.Label(schedule_frame, text="Default Day:").grid(row=1, column=0, sticky=self.tk.W, padx=5, pady=5)

        days = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
        self.ttk.Combobox(schedule_frame, textvariable=self.default_day_var,
                          values=days, state="readonly", width=15).grid(row=1, column=1, sticky=self.tk.W, padx=5,
                                                                        pady=5)

        # Output directory
        self.ttk.Label(schedule_frame, text="Output Directory:").grid(row=2, column=0, sticky=self.tk.W, padx=5, pady=5)

        output_frame = self.ttk.Frame(schedule_frame)
        output_frame.grid(row=2, column=1, sticky=self.tk.W, padx=5, pady=5)

        self.ttk.Entry(output_frame, textvariable=self.output_dir_var, width=30).pack(side=self.tk.LEFT)
        self.ttk.Button(output_frame, text="Browse...", command=self._browse_output_dir).pack(side=self.tk.LEFT,
                                                                                              padx=(5, 0))

        # Save button
        save_schedule_frame = self.ttk.Frame(schedule_frame)
        save_schedule_frame.grid(row=3, column=0, columnspan=2, sticky=self.tk.E, padx=5, pady=(15, 5))

        self.ttk.Button(save_schedule_frame, text="Save Schedule Config",
                        command=self._save_schedule_config).pack(side=self.tk.RIGHT)

        # Excel configuration tab
        excel_tab = self.ttk.Frame(notebook)
        notebook.add(excel_tab, text="Excel Settings")

        excel_frame = self.ttk.LabelFrame(excel_tab, text="Excel Processing Configuration")
        excel_frame.pack(fill=self.tk.BOTH, expand=True, padx=10, pady=10)

        # Excel visibility option
        self.excel_visible_var = self.tk.BooleanVar(value=False)
        self.ttk.Checkbutton(
            excel_frame,
            text="Show Excel During Processing (Debug Mode)",
            variable=self.excel_visible_var
        ).grid(row=0, column=0, columnspan=2, sticky=self.tk.W, padx=5, pady=5)

        # Cleanup option
        self.excel_cleanup_var = self.tk.BooleanVar(value=True)
        self.ttk.Checkbutton(
            excel_frame,
            text="Cleanup Excel After Job Completion",
            variable=self.excel_cleanup_var
        ).grid(row=1, column=0, columnspan=2, sticky=self.tk.W, padx=5, pady=5)

        # Max retries
        self.ttk.Label(excel_frame, text="Max Retries:").grid(row=2, column=0, sticky=self.tk.W, padx=5, pady=5)
        self.excel_retries_var = self.tk.StringVar(value="3")
        self.ttk.Entry(excel_frame, textvariable=self.excel_retries_var, width=5).grid(
            row=2, column=1, sticky=self.tk.W, padx=5, pady=5
        )

        # Save button
        excel_btn_frame = self.ttk.Frame(excel_frame)
        excel_btn_frame.grid(row=3, column=0, columnspan=2, sticky=self.tk.E, padx=5, pady=(15, 5))

        self.ttk.Button(
            excel_btn_frame,
            text="Save Excel Config",
            command=self._save_excel_config
        ).pack(side=self.tk.RIGHT)

        # Force cleanup button
        self.ttk.Button(
            excel_btn_frame,
            text="Force Excel Cleanup",
            command=self._force_excel_cleanup
        ).pack(side=self.tk.RIGHT, padx=(0, 10))

    def _save_excel_config(self):
        """Save Excel configuration"""
        # Create Excel configuration
        excel_config = {
            "visible": self.excel_visible_var.get(),
            "cleanup_after_job": self.excel_cleanup_var.get(),
            "max_retries": int(self.excel_retries_var.get())
        }

        # Get current config
        scheduler_config = self.scheduler.scheduler_config
        scheduler_config["excel"] = excel_config

        # Save entire configuration
        if self.scheduler.save_scheduler_config():
            self.messagebox.showinfo("Success", "Excel configuration saved")
        else:
            self.messagebox.showerror("Error", "Failed to save Excel configuration")

    def _force_excel_cleanup(self):
        """Force cleanup of Excel processes"""
        try:
            ensure_excel_closed()
            self.messagebox.showinfo("Success", "Excel processes cleaned up")
        except Exception as e:
            self.messagebox.showerror("Error", f"Failed to clean up Excel processes: {e}")

    def _load_config(self):
        """Load configuration from scheduler"""
        # Load email configuration
        email_config = self.scheduler.get_email_config()
        self.email_enabled_var.set(email_config.get("enabled", False))
        self.smtp_server_var.set(email_config.get("smtp_server", ""))
        self.smtp_port_var.set(str(email_config.get("smtp_port", 587)))
        self.use_tls_var.set(email_config.get("use_tls", True))
        self.username_var.set(email_config.get("username", ""))
        self.password_var.set(email_config.get("password", ""))
        self.from_address_var.set(email_config.get("from_address", ""))
        self.admin_address_var.set(email_config.get("admin_address", ""))

        # Update email config state
        self._update_email_config_state()

        # Load schedule configuration
        schedule_config = self.scheduler.get_schedule_config()
        self.default_time_var.set(schedule_config.get("default_time", "08:00"))
        self.default_day_var.set(schedule_config.get("default_day", "monday"))
        self.output_dir_var.set(schedule_config.get("output_dir", "automated_output"))

        # Load Excel configuration
        excel_config = self.scheduler.scheduler_config.get("excel", {})
        self.excel_visible_var.set(excel_config.get("visible", False))
        self.excel_cleanup_var.set(excel_config.get("cleanup_after_job", True))
        self.excel_retries_var.set(str(excel_config.get("max_retries", 3)))

        # Load jobs
        self._refresh_job_list()

    def _refresh_job_list(self):
        """Refresh the job list"""
        # Clear existing items
        for item in self.job_tree.get_children():
            self.job_tree.delete(item)

        # Get jobs from scheduler
        jobs = self.scheduler.get_jobs()

        # Add jobs to treeview
        for job in jobs:
            job_id = job.get("job_id")
            analytics_id = job.get("analytics_id")

            # Format schedule
            schedule_type = job.get("schedule_type", "daily")
            schedule_time = job.get("schedule_time", "08:00")
            schedule_day = job.get("schedule_day", "")

            if schedule_type == "daily":
                schedule = f"Daily at {schedule_time}"
            elif schedule_type == "weekly":
                schedule = f"Weekly on {schedule_day.capitalize()} at {schedule_time}"
            elif schedule_type == "monthly":
                schedule = f"Monthly on day {schedule_day} at {schedule_time}"
            else:
                schedule = "Unknown"

            # Email status
            email = "Yes" if job.get("send_email", False) else "No"

            self.job_tree.insert("", self.tk.END, values=(job_id, analytics_id, schedule, email))

    def _update_scheduler_status(self):
        """Update the scheduler status display"""
        if self.scheduler.is_running():
            self.status_label.config(text="Scheduler is running")
            self.start_btn.config(state=self.tk.DISABLED)
            self.stop_btn.config(state=self.tk.NORMAL)
        else:
            self.status_label.config(text="Scheduler is not running")
            self.start_btn.config(state=self.tk.NORMAL)
            self.stop_btn.config(state=self.tk.DISABLED)

    def _on_job_selected(self, event):
        """Handle job selection event"""
        selection = self.job_tree.selection()
        if not selection:
            return

        # Get the selected job ID
        values = self.job_tree.item(selection[0], "values")
        job_id = values[0]

        # Find job in configuration
        jobs = self.scheduler.get_jobs()
        job = None

        for j in jobs:
            if j.get("job_id") == job_id:
                job = j
                break

        if not job:
            return

        # Update form fields
        self.job_id_var.set(job.get("job_id", ""))
        self.analytics_id_var.set(job.get("analytics_id", ""))
        self.schedule_type_var.set(job.get("schedule_type", "daily"))
        self.schedule_time_var.set(job.get("schedule_time", "08:00"))
        self.schedule_day_var.set(job.get("schedule_day", "monday"))
        self.data_source_var.set(job.get("data_source_pattern", ""))
        self.send_email_var.set(job.get("send_email", False))
        self.generate_individual_var.set(job.get("generate_individual_reports", False))

        # Update recipients
        self.recipients = job.get("email_recipients", []).copy()
        self._refresh_recipients_list()

        # Update UI based on selections
        self._update_schedule_options()
        self._update_email_options()

    def _new_job(self):
        """Create a new job"""
        # Clear form fields
        import datetime
        self.job_id_var.set(f"job_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}")
        self.analytics_id_var.set("")
        self.schedule_type_var.set("daily")
        self.schedule_time_var.set("08:00")
        self.schedule_day_var.set("monday")
        self.data_source_var.set("")
        self.send_email_var.set(False)
        self.generate_individual_var.set(False)

        # Clear recipients
        self.recipients = []
        self._refresh_recipients_list()

        # Update UI based on selections
        self._update_schedule_options()
        self._update_email_options()

    def _delete_job(self):
        """Delete the selected job"""
        selection = self.job_tree.selection()
        if not selection:
            self.messagebox.showinfo("Info", "Please select a job to delete")
            return

        # Get the selected job ID
        values = self.job_tree.item(selection[0], "values")
        job_id = values[0]

        # Confirm deletion
        if not self.messagebox.askyesno("Confirm", f"Are you sure you want to delete job '{job_id}'?"):
            return

        # Delete job
        if self.scheduler.remove_job(job_id):
            self.messagebox.showinfo("Success", f"Job '{job_id}' deleted")
            self._refresh_job_list()
        else:
            self.messagebox.showerror("Error", f"Failed to delete job '{job_id}'")

    def _run_job_now(self):
        """Run the selected job now"""
        selection = self.job_tree.selection()
        if not selection:
            self.messagebox.showinfo("Info", "Please select a job to run")
            return

        # Get the selected job ID
        values = self.job_tree.item(selection[0], "values")
        job_id = values[0]

        # Run job
        if self.scheduler.run_job_now(job_id):
            self.messagebox.showinfo("Success", f"Job '{job_id}' started")
        else:
            self.messagebox.showerror("Error", f"Failed to run job '{job_id}'")

    def _save_job(self):
        """Save the current job configuration"""
        # Validate required fields
        job_id = self.job_id_var.get().strip()
        analytics_id = self.analytics_id_var.get().strip()
        data_source_pattern = self.data_source_var.get().strip()

        if not job_id:
            self.messagebox.showerror("Error", "Job ID is required")
            return

        if not analytics_id:
            self.messagebox.showerror("Error", "Analytics ID is required")
            return

        if not data_source_pattern:
            self.messagebox.showerror("Error", "Data file pattern is required")
            return

        # Create job configuration
        job_config = {
            "job_id": job_id,
            "analytics_id": analytics_id,
            "schedule_type": self.schedule_type_var.get(),
            "schedule_time": self.schedule_time_var.get(),
            "data_source_pattern": data_source_pattern,
            "send_email": self.send_email_var.get(),
            "generate_individual_reports": self.generate_individual_var.get()
        }

        # Add day for weekly and monthly schedules
        if self.schedule_type_var.get() in ["weekly", "monthly"]:
            job_config["schedule_day"] = self.schedule_day_var.get()

        # Add recipients if email is enabled
        if self.send_email_var.get():
            job_config["email_recipients"] = self.recipients.copy()

        # Save job
        if self.scheduler.add_job(job_config):
            self.messagebox.showinfo("Success", f"Job '{job_id}' saved")
            self._refresh_job_list()
        else:
            self.messagebox.showerror("Error", f"Failed to save job '{job_id}'")

    def _update_schedule_options(self):
        """Update schedule options based on schedule type"""
        schedule_type = self.schedule_type_var.get()

        if schedule_type == "daily":
            self.day_label.grid_remove()
            self.day_frame.grid_remove()
        else:
            self.day_label.grid()
            self.day_frame.grid()

            if schedule_type == "weekly":
                self.week_frame.pack(fill=self.tk.X)
                self.month_frame.pack_forget()
            else:  # monthly
                self.week_frame.pack_forget()
                self.month_frame.pack(fill=self.tk.X)

    def _update_email_options(self):
        """Update email options based on send_email checkbox"""
        if self.send_email_var.get():
            self.recipients_frame.grid()
        else:
            self.recipients_frame.grid_remove()

    def _update_email_config_state(self):
        """Update email configuration field states based on enabled checkbox"""
        state = self.tk.NORMAL if self.email_enabled_var.get() else self.tk.DISABLED

        self.smtp_server_entry.config(state=state)
        self.smtp_port_entry.config(state=state)
        self.use_tls_check.config(state=state)
        self.username_entry.config(state=state)
        self.password_entry.config(state=state)
        self.from_address_entry.config(state=state)
        self.admin_address_entry.config(state=state)
        self.test_email_btn.config(state=state)

    def _add_recipient(self):
        """Add a recipient to the list"""
        recipient = self.recipient_var.get().strip()

        if not recipient:
            return

        # Simple email validation
        if "@" not in recipient:
            self.messagebox.showerror("Error", "Invalid email address")
            return

        # Add to list if not already present
        if recipient not in self.recipients:
            self.recipients.append(recipient)
            self._refresh_recipients_list()
            self.recipient_var.set("")

    def _remove_recipient(self):
        """Remove a recipient from the list"""
        selection = self.recipients_list.curselection()
        if not selection:
            return

        # Get selected index
        index = selection[0]

        # Remove from list
        del self.recipients[index]
        self._refresh_recipients_list()

    def _refresh_recipients_list(self):
        """Refresh the recipients list"""
        self.recipients_list.delete(0, self.tk.END)

        for recipient in self.recipients:
            self.recipients_list.insert(self.tk.END, recipient)

    def _browse_data_pattern(self):
        """Browse for data file pattern"""
        # First, ask for directory
        directory = self.filedialog.askdirectory(title="Select Data Directory")
        if not directory:
            return

        # Then, ask for filename pattern
        self.data_source_var.set(os.path.join(directory, "*.xlsx"))

    def _browse_output_dir(self):
        """Browse for output directory"""
        directory = self.filedialog.askdirectory(title="Select Output Directory")
        if directory:
            self.output_dir_var.set(directory)

    def _save_email_config(self):
        """Save email configuration"""
        # Create email configuration
        email_config = {
            "enabled": self.email_enabled_var.get(),
            "smtp_server": self.smtp_server_var.get(),
            "smtp_port": int(self.smtp_port_var.get()),
            "use_tls": self.use_tls_var.get(),
            "username": self.username_var.get(),
            "password": self.password_var.get(),
            "from_address": self.from_address_var.get(),
            "admin_address": self.admin_address_var.get()
        }

        # Save configuration
        if self.scheduler.update_email_config(email_config):
            self.messagebox.showinfo("Success", "Email configuration saved")
        else:
            self.messagebox.showerror("Error", "Failed to save email configuration")

    def _save_schedule_config(self):
        """Save schedule configuration"""
        # Create schedule configuration
        schedule_config = {
            "default_time": self.default_time_var.get(),
            "default_day": self.default_day_var.get(),
            "output_dir": self.output_dir_var.get()
        }

        # Save configuration
        if self.scheduler.update_schedule_config(schedule_config):
            self.messagebox.showinfo("Success", "Schedule configuration saved")
        else:
            self.messagebox.showerror("Error", "Failed to save schedule configuration")

    def _start_scheduler(self):
        """Start the scheduler"""
        if self.scheduler.start_scheduler():
            self.messagebox.showinfo("Success", "Scheduler started")
            self._update_scheduler_status()
        else:
            self.messagebox.showerror("Error", "Failed to start scheduler")

    def _stop_scheduler(self):
        """Stop the scheduler"""
        if self.scheduler.stop_scheduler():
            self.messagebox.showinfo("Success", "Scheduler stopped")
            self._update_scheduler_status()
        else:
            self.messagebox.showerror("Error", "Failed to stop scheduler")

    def _test_email(self):
        """Test email configuration"""
        # Get recipient
        recipient = self.admin_address_var.get()
        if not recipient:
            recipient = self.messagebox.askstring("Recipient", "Enter email address for test:")
            if not recipient:
                return

        # Create email configuration
        email_config = {
            "enabled": self.email_enabled_var.get(),
            "smtp_server": self.smtp_server_var.get(),
            "smtp_port": int(self.smtp_port_var.get()),
            "use_tls": self.use_tls_var.get(),
            "username": self.username_var.get(),
            "password": self.password_var.get(),
            "from_address": self.from_address_var.get(),
            "admin_address": self.admin_address_var.get()
        }

        # Update config (temporarily)
        self.scheduler.update_email_config(email_config)

        # Test email
        success, message = self.scheduler.test_email_configuration(recipient)

        if success:
            self.messagebox.showinfo("Success", message)
        else:
            self.messagebox.showerror("Error", message)