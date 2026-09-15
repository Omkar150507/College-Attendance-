# Production settings for simultaneous college users.
bind = "0.0.0.0:10000"
workers = 2
threads = 4
timeout = 60
keepalive = 5
worker_tmp_dir = "/dev/shm"
preload_app = False
