import { cp, mkdir } from 'node:fs/promises'

await mkdir('dist/server', { recursive: true })
await cp('dist/testpilot_ai_web/index.js', 'dist/server/index.js')
await cp('dist/testpilot_ai_web/wrangler.json', 'dist/server/wrangler.json')
