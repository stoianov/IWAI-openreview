from collections import defaultdict
import time

USE_INSTITUTION_CONFLICTS = True
MAX_COAUTHOR_YEARS = 5
SLEEP_BETWEEN_PROFILE_CALLS = 0.05

# ============================================================
# PROFILE CACHE
# ============================================================
def get_profiles(reviewer_ids):
    print("\nDownloading reviewer profiles...")
    profile_cache = {}
    for rid in tqdm(reviewer_ids):
        try:
            profile = client.get_profile(rid)
            profile_cache[rid] = profile
        except Exception as e:
            print(f"Could not retrieve profile for {rid}")
            profile_cache[rid] = None
        time.sleep(SLEEP_BETWEEN_PROFILE_CALLS)
