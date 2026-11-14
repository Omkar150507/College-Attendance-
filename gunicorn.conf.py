import os

# Render supplies PORT at runtime. This also keeps local fallback compatibility.
bind = f"0.0.0.0:{os.environ.get('PORT', '10000')}"
# Render's free instance has little memory: it sets WEB_CONCURRENCY=1 by default, which we respect.
workers = int(os.environ.get("WEB_CONCURRENCY", "1"))
threads = 4
timeout = 60
keepalive = 5
worker_tmp_dir = "/dev/shm"
preload_app = False
