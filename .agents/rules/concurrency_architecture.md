# High-Concurrency AI Generation & Queue Architecture Rules

## AI Generation Rate-Limits & Concurrency Limits
1. **Hugging Face Inference API**:
   - Concurrency limit: ~2–5 parallel requests per token.
   - Rate limit: ~50–200 requests per hour per account.
2. **Google Gemini API**:
   - Free tier limit: 15 Requests Per Minute (RPM) / 1,500 Requests Per Day (RPD).
3. **Multi-Teacher Parallel Handling Pattern**:
   - When multiple teachers generate AI papers concurrently, use an **Asynchronous Task Queue** pattern:
     1. Accept request instantly (`HTTP 202 Accepted`), return `job_id` and queue position.
     2. Process queue sequentially (1 request at a time) via a background worker thread.
     3. Insert notification into `teacher_notifications` table upon completion.
     4. Notify teacher via dashboard notification bell.
