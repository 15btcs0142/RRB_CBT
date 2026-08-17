# Redis Setup & Installation Guide for Windows

This guide provides options for running Redis on Windows for RRB CBT background job queueing (`redis` + `rq`).

---

## Option 1: Native Windows Redis via Memurai (Recommended for Windows)

Memurai is a 100% Redis-compatible, high-performance database engine built specifically for Windows.

1. Download Memurai Developer Edition (Free):
   [https://www.memurai.com/get-memurai](https://www.memurai.com/get-memurai)
2. Run the `.msi` installer. It will automatically start Memurai as a Windows Service listening on port `6379`.
3. Verify connection:
   ```cmd
   python -c "import redis; r=redis.Redis(); print(r.ping())"
   ```
   *Expected Output: `True`*

---

## Option 2: Redis for Windows (MSI Installer)

1. Download Redis 5.0 MSI installer for Windows:
   [https://github.com/tporadowski/redis/releases](https://github.com/tporadowski/redis/releases)
2. Run `Redis-x64-5.0.14.1.msi` installer.
3. Keep default settings (`Port 6379`, `Add to PATH`).
4. Redis will automatically start as a Windows Service.

---

## Option 3: Redis via WSL2 (Windows Subsystem for Linux)

If you have WSL2 (Ubuntu) enabled:

1. Open WSL terminal:
   ```bash
   sudo apt update
   sudo apt install redis-server -y
   sudo service redis-server start
   ```
2. Test connection from Windows CMD/PowerShell:
   ```cmd
   redis-cli ping
   ```

---

## Running the RQ Worker

Once Redis is running:

1. Open a terminal / command prompt in the RRB CBT project folder.
2. Launch the RQ worker process:
   ```cmd
   python rq_worker.py
   ```
3. Or launch using option `[11]` in `RRB_CBT_manager.bat`.

---

## Automatic Fallback Guarantee

If Redis is **not installed or stopped**, `app.py` automatically detects that Redis is offline and **reverts to the built-in Thread Queue worker (`queue.Queue()`)**. Paper generation, `.doc` file saving, and teacher notifications will continue working without interruption.
