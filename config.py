"""
config.py
----------
Central configuration for the Hospital Review Automation Bot.

Every configurable value in the project should live here.
No hardcoded values should exist elsewhere in the codebase.
"""

from pathlib import Path

# =============================================================================
# PROJECT PATHS
# =============================================================================

PROJECT_ROOT = Path(__file__).parent

OUTPUT_DIR = PROJECT_ROOT / "output"
LOG_DIR = PROJECT_ROOT / "logs"
CACHE_DIR = PROJECT_ROOT / "cache"
SCREENSHOT_DIR = PROJECT_ROOT / "screenshots"

OUTPUT_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)
CACHE_DIR.mkdir(exist_ok=True)
SCREENSHOT_DIR.mkdir(exist_ok=True)

# =============================================================================
# WEBSITE CONFIGURATION
# =============================================================================

BASE_URL = "https://karunahealthlifepartner.com/index.php/account/admin"

USERNAME = "karuna__admin"

PASSWORD = "karuna__admin"

HEADLESS = False

DEFAULT_TIMEOUT = 30000

# =============================================================================
# REVIEW SETTINGS
# =============================================================================

MAX_REVIEWS_PER_HOSPITAL = 5

LOW_REVIEW_THRESHOLD = 3

PENDING_REVIEW = True

# =============================================================================
# SEARCH SETTINGS
# =============================================================================

SEARCH_ENGINE = "duckduckgo"

SEARCH_DELAY_MIN = 2

SEARCH_DELAY_MAX = 5

# =============================================================================
# FILES
# =============================================================================

HOSPITAL_LIST_FILE = OUTPUT_DIR / "hospital_names.xlsx"

REVIEWS_FILE = OUTPUT_DIR / "reviews.xlsx"

NO_REVIEWS_DOC = OUTPUT_DIR / "no_reviews.docx"

LOW_REVIEWS_DOC = OUTPUT_DIR / "low_reviews.docx"

AUTOMATION_LOG = LOG_DIR / "automation.log"

CHECKPOINT_FILE = CACHE_DIR / "checkpoint.json"

PROCESSED_FILE = CACHE_DIR / "processed_hospitals.txt"

GOOGLE_CACHE_FILE = CACHE_DIR / "review_cache.json"

SKIP_LIST_FILE = PROJECT_ROOT / "hospitals_no_photo.docx"

# =============================================================================
# DEVELOPMENT
# =============================================================================

TEST_MODE = True

TEST_HOSPITAL = ""

SAVE_SCREENSHOTS = True

VERBOSE = True