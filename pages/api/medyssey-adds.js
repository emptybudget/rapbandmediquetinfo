import { redis } from '../../lib/redis';

const ADDS_KEY = 'rapband:medysseyAdds';

export default async function handler(req, res) {
  if (req.method === 'GET') {
    if (!redis) return res.json({});
    try {
      const data = await redis.hgetall(ADDS_KEY);
      return res.json(data || {});
    } catch {
      return res.json({});
    }
  }

  if (req.method === 'POST') {
    const { hospital, instrument, size } = req.body || {};
    if (!hospital || !instrument) {
      return res.status(400).json({ error: 'hospital and instrument required' });
    }
    const key = `${hospital}|${instrument}|${size ?? ''}`;
    if (redis) {
      try {
        await redis.hset(ADDS_KEY, { [key]: '1' });
      } catch {}
    }
    return res.json({ ok: true, key });
  }

  res.status(405).end();
}
