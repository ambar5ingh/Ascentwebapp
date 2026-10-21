// Password gate. The app is a static site, so this only keeps casual visitors out:
// the check runs in the browser against a SHA-256 hash of the password.
const HASH = 'c565b3bccf5fbec42a15be8a31c67d9eba7a7cf06a81da9a8f41688cda806de9';
const KEY = 'ascent-auth';

export function sha256(str) {
  const K = [0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5, 0xd807aa98, 0x12835b01,
    0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174, 0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f,
    0x4a7484aa, 0x5cb0a9dc, 0x76f988da, 0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
    0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85, 0xa2bfe8a1, 0xa81a664b, 0xc24b8b70,
    0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070, 0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a,
    0x5b9cca4f, 0x682e6ff3, 0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2];
  const bytes = new TextEncoder().encode(str);
  const l = bytes.length;
  const n = (((l + 9) >> 6) + 1) << 4;
  const w = new Uint32Array(n);
  for (let i = 0; i < l; i++) w[i >> 2] |= bytes[i] << (24 - (i % 4) * 8);
  w[l >> 2] |= 0x80 << (24 - (l % 4) * 8);
  w[n - 1] = l * 8;
  const H = [0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19];
  const W = new Uint32Array(64);
  const rot = (x, k) => (x >>> k) | (x << (32 - k));
  for (let b = 0; b < n; b += 16) {
    for (let t = 0; t < 64; t++) {
      if (t < 16) W[t] = w[b + t];
      else {
        const s0 = rot(W[t - 15], 7) ^ rot(W[t - 15], 18) ^ (W[t - 15] >>> 3);
        const s1 = rot(W[t - 2], 17) ^ rot(W[t - 2], 19) ^ (W[t - 2] >>> 10);
        W[t] = (W[t - 16] + s0 + W[t - 7] + s1) | 0;
      }
    }
    let [a, b2, c, d, e, f, g, h] = H;
    for (let t = 0; t < 64; t++) {
      const S1 = rot(e, 6) ^ rot(e, 11) ^ rot(e, 25);
      const ch = (e & f) ^ (~e & g);
      const t1 = (h + S1 + ch + K[t] + W[t]) | 0;
      const S0 = rot(a, 2) ^ rot(a, 13) ^ rot(a, 22);
      const maj = (a & b2) ^ (a & c) ^ (b2 & c);
      const t2 = (S0 + maj) | 0;
      h = g; g = f; f = e; e = (d + t1) | 0; d = c; c = b2; b2 = a; a = (t1 + t2) | 0;
    }
    H[0] = (H[0] + a) | 0; H[1] = (H[1] + b2) | 0; H[2] = (H[2] + c) | 0; H[3] = (H[3] + d) | 0;
    H[4] = (H[4] + e) | 0; H[5] = (H[5] + f) | 0; H[6] = (H[6] + g) | 0; H[7] = (H[7] + h) | 0;
  }
  return H.map((x) => (x >>> 0).toString(16).padStart(8, '0')).join('');
}

function stored() { try { return sessionStorage.getItem(KEY); } catch (e) { return null; } }

export function isAuthed() { return stored() === HASH; }
export function logout() { try { sessionStorage.removeItem(KEY); } catch (e) { /* ignore */ } }

// Show the login screen until the right password is entered.
export function requireLogin(root) {
  if (isAuthed()) { root.remove(); return Promise.resolve(); }
  return new Promise((resolve) => {
    root.style.display = 'flex';
    const form = root.querySelector('form');
    const input = root.querySelector('input');
    const msg = root.querySelector('.login-error');
    input.focus();
    form.addEventListener('submit', (e) => {
      e.preventDefault();
      if (sha256(input.value) === HASH) {
        try { sessionStorage.setItem(KEY, HASH); } catch (err) { /* private mode */ }
        root.remove();
        resolve();
      } else {
        msg.textContent = 'Incorrect password. Please try again.';
        input.select();
      }
    });
  });
}
