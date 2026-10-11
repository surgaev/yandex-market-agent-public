"""
Determine which offers are currently live on the storefront.

Uses Partner API: POST /campaigns/{campaignId}/offers
Offer statuses:
  PUBLISHED              — на витрине, доступен покупателям
  NO_STOCKS              — нет остатков
  DISABLED_AUTOMATICALLY — автоматически отключён
  ARCHIVED               — в архиве
  и другие

Товар считается «на витрине» если status == PUBLISHED.
"""

import time
from .config import OAUTH_TOKEN, CAMPAIGN_ID, YM_API_BASE
from .yandex_client import _request_raw_with_retry


def get_storefront_skus() -> set[str]:
    """
    Returns set of shopSku IDs that are currently PUBLISHED (live on storefront).
    Falls back to empty set on error so the agent still runs.
    """
    try:
        headers = {"Api-Key": OAUTH_TOKEN, "Content-Type": "application/json"}
        published: set[str] = set()
        page_token = None
        seen_tokens: set[str] = set()

        while True:
            params = {"pageToken": page_token} if page_token else {}
            r = _request_raw_with_retry(
                "POST", f"{YM_API_BASE}/campaigns/{CAMPAIGN_ID}/offers",
                headers=headers, json={}, params=params,
            )
            r.raise_for_status()
            data = r.json()

            for offer in data.get("result", {}).get("offers", []):
                if offer.get("status") == "PUBLISHED":
                    published.add(offer["offerId"])

            page_token = data.get("result", {}).get("paging", {}).get("nextPageToken")
            if not page_token or page_token in seen_tokens:
                break
            seen_tokens.add(page_token)
            time.sleep(0.2)

        return published

    except Exception as e:
        print(f"  [warn] storefront API failed: {e}")
        return set()
