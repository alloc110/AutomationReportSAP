import os
import time
import logging
import schedule
from dotenv import load_dotenv
from automation import run_automation

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
logger = logging.getLogger("Scheduler")

SCHEDULE_TIME = os.getenv("RUN_AT_TIME", "08:00")
RUN_ON_STARTUP = os.getenv("RUN_ON_STARTUP", "false").lower() in ("true", "1", "yes")

def job():
    logger.info("Triggering scheduled job...")
    try:
        run_automation()
    except Exception as e:
        logger.error(f"Job execution failed: {e}")

def main():
    logger.info(f"Scheduler service started. Daily execution scheduled at: {SCHEDULE_TIME}")

    if RUN_ON_STARTUP:
        logger.info("Running job immediately on startup...")
        job()

    schedule.every().day.at(SCHEDULE_TIME).do(job)

    while True:
        schedule.run_pending()
        time.sleep(30)

if __name__ == "__main__":
    main()
