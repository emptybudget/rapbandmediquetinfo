import { redis, LOG_KEY, MAX_LOG } from '../../lib/redis';

export default async function handler(req, res) {
  if (!redis) return res.status(503).json({ error: 'Redis not configured' });

  if (req.method === 'GET') {
    try {
      const raw = await redis.lrange(LOG_KEY, 0, MAX_LOG - 1);
      const entries = raw.map(e => (typeof e === 'string' ? JSON.parse(e) : e));
      return res.json(entries);
    } catch {
      return res.json([]);
    }
  }

  if (req.method === 'POST') {
    try {
      await redis.lpush(LOG_KEY, JSON.stringify(req.body));
      await redis.ltrim(LOG_KEY, 0, MAX_LOG - 1);
    } catch {}
    return res.json({ ok: true });
  }

  if (req.method === 'DELETE') {
    try { await redis.del(LOG_KEY); } catch {}
    return res.json({ ok: true });
  }

  res.status(405).end();
}
