# Security Documentation

## Patches Applied (v4 → Production)

### 1. Password Hashing (CRITICAL)
- **Before:** Plaintext passwords stored in source (`"manager123"`)
- **After:** SHA-256 hashes of `"username:password"` using Web Crypto API
- **How login works:** User input → `crypto.subtle.digest("SHA-256", "user:pass")` → compare with stored hash
- **To change a password:** Compute `SHA-256("username:newpassword")` at https://emn178.github.io/online-tools/sha256.html and update `CREDENTIALS` in `App.jsx`

### 2. Rate Limiting
- Max 5 login attempts per username per 60 seconds
- Implemented client-side via `loginAttempts` map with timestamp-based reset
- **Production note:** For stronger protection, add server-side rate limiting (Vercel Edge Functions, Cloudflare Workers)

### 3. File Upload Validation
- Maximum file size: 50 MB
- Extension whitelist: `.xlsx .xls .xlsm .xlsb .xltx .xltm .xlt .csv .ods .tsv`
- Files are parsed as binary (ArrayBuffer), never executed

### 4. Export Filename Sanitization
- Download filenames stripped of special characters: `[^a-zA-Z0-9_-]` → `_`
- Maximum filename length: 50 characters

### 5. Security Headers (via vercel.json / netlify.toml)
| Header | Value |
|--------|-------|
| X-Frame-Options | DENY |
| X-Content-Type-Options | nosniff |
| X-XSS-Protection | 1; mode=block |
| Referrer-Policy | strict-origin-when-cross-origin |
| Strict-Transport-Security | max-age=31536000; includeSubDomains |
| Permissions-Policy | camera=(), microphone=(), geolocation=() |

### 6. Content Security Policy (index.html meta tag)
- `default-src 'self'` — blocks external resource loading
- `frame-ancestors 'none'` — prevents clickjacking
- `base-uri 'self'` — prevents base tag injection
- `connect-src 'self'` — blocks unauthorized XHR/fetch calls

## False Positives (Verified Safe)
- **ReDoS patterns:** String concatenation `+` operators, not regex patterns
- **HTTP URL:** W3C SVG namespace URI (`http://www.w3.org/2000/svg`) — not a network request
- **JSON.parse:** Used only on our own serialized objects (deep copy pattern)
- **Math.random():** Used only for UI element IDs, not cryptographic purposes

## Known Limitations (Frontend-Only App)
Since this is a client-side React app with no backend:
- Credentials are validated in the browser — determined attackers can inspect source
- **Recommendation for production:** Replace CREDENTIALS with a real auth backend (Firebase Auth, Auth0, or your company's SSO/LDAP)
- Data is held in React state — refreshing the page loses imported data
- **Recommendation:** Add localStorage persistence or a backend API

## Dependency License Summary
| Package | Version | License | Status |
|---------|---------|---------|--------|
| react | 18.x | MIT | ✅ Free for any use |
| react-dom | 18.x | MIT | ✅ Free for any use |
| recharts | 2.x | MIT | ✅ Free for any use |
| vite | 5.x | MIT | ✅ Free for any use |
| @vitejs/plugin-react | 4.x | MIT | ✅ Free for any use |
| xlsx | **0.18.5** | Apache-2.0 | ✅ Free for any use (pinned) |

**⚠️ xlsx WARNING:** Do not run `npm update` or change `"xlsx": "0.18.5"` to a range like `"^0.18.5"`.
Versions 0.19.0 and later are under a commercial license.
