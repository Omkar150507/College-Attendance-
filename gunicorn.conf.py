import os

# Render provides PORT at runtime. Keep 10000 only as a local fallback.
bind = f"0.0.0.0:{os.environ.get('PORT', '10000')}"
workers = 2
threads = 4
timeout = 60
keepalive = 5
worker_tmp_dir = "/dev/shm"
preload_app = False
