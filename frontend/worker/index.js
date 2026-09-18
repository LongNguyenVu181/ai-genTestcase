export default {
  async fetch(request, env) {
    const url = new URL(request.url)
    if (url.pathname === '/api/model') {
      if (request.method !== 'POST') return new Response('Method not allowed', { status: 405 })
      let payload
      try {
        payload = await request.json()
      } catch {
        return new Response(JSON.stringify({ message: 'Payload AI không hợp lệ.' }), { status: 400, headers: { 'Content-Type': 'application/json' } })
      }
      let providerUrl
      try {
        providerUrl = new URL(String(payload.baseUrl || ''))
      } catch {
        return new Response(JSON.stringify({ message: 'Base URL AI không hợp lệ.' }), { status: 400, headers: { 'Content-Type': 'application/json' } })
      }
      if (providerUrl.protocol !== 'https:' || !providerUrl.hostname.endsWith('.maas.aliyuncs.com')) {
        return new Response(JSON.stringify({ message: 'Bản deploy này chỉ hỗ trợ Qwen/DashScope (maas.aliyuncs.com).' }), { status: 400, headers: { 'Content-Type': 'application/json' } })
      }
      const apiKey = String(payload.apiKey || '').trim()
      const model = String(payload.model || '').trim()
      const prompt = String(payload.prompt || '')
      if (!apiKey || !model || !prompt) {
        return new Response(JSON.stringify({ message: 'Thiếu API key, model hoặc nội dung yêu cầu.' }), { status: 400, headers: { 'Content-Type': 'application/json' } })
      }
      const upstream = await fetch(`${providerUrl.toString().replace(/\/+$/, '')}/chat/completions`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${apiKey}` },
        body: JSON.stringify({
          model,
          messages: [{ role: 'user', content: prompt }],
          temperature: 0.1,
        }),
      })
      return new Response(upstream.body, {
        status: upstream.status,
        headers: { 'Content-Type': upstream.headers.get('Content-Type') || 'application/json', 'Cache-Control': 'no-store' },
      })
    }
    const asset = await env.ASSETS.fetch(request)
    if (asset.status !== 404 || url.pathname.includes('.')) return asset
    return env.ASSETS.fetch(new Request(new URL('/index.html', request.url)))
  },
}
