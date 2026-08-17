import os
import sys
from dotenv import load_dotenv

# Load environment variables
load_dotenv('.env')

import redis
from rq import Worker, Queue, Connection

listen = ['paper_generation', 'default']
redis_host = os.environ.get('REDIS_HOST', 'localhost')
redis_port = int(os.environ.get('REDIS_PORT', 6379))
redis_db = int(os.environ.get('REDIS_DB', 0))

def start_worker():
    print(f"[*] Starting RRB CBT RQ Worker listening on Redis {redis_host}:{redis_port} (Queues: {', '.join(listen)})...")
    try:
        conn = redis.Redis(host=redis_host, port=redis_port, db=redis_db)
        conn.ping()
        print("[+] Redis connection successful! Worker ready for jobs.")
        with Connection(conn):
            worker = Worker(list(map(Queue, listen)))
            worker.work()
    except Exception as e:
        print(f"[!] Unable to connect to Redis at {redis_host}:{redis_port}: {e}")
        print("[!] Make sure Redis server (Memurai / Redis Windows / WSL) is running.")
        sys.exit(1)

if __name__ == '__main__':
    start_worker()
