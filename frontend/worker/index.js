export default {
  async fetch(request, env) {
    const asset = await env.ASSETS.fetch(request)
    if (asset.status !== 404 || new URL(request.url).pathname.includes('.')) return asset
    return env.ASSETS.fetch(new Request(new URL('/index.html', request.url)))
  },
}
