import ntplib
from datetime import datetime, timezone, timedelta
from zoneinfo import ZoneInfo 

def get_internet_time():
    # client = ntplib.NTPClient()
    # response = client.request('time.google.com')
    #dt=datetime.fromtimestamp(response.tx_time, tz=timezone.utc) + timedelta(hours=5, minutes=30)
    dt=datetime.now(ZoneInfo("Asia/Kolkata"))
    return dt.strftime(f"{dt.day}/{dt.month}/{dt.year} {dt.hour}:{dt.minute:02d}")

# Usage
# print("Accurate Internet Time (UTC):", get_internet_time())