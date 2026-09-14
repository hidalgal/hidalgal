'use strict';

const assert = require('node:assert');
const { encryptToken, decryptToken, getRandomHexString } = require('./aesCtrToken');

function run(name, fn) {
  try {
    fn();
    console.log(`OK   - ${name}`);
  } catch (err) {
    console.error(`FAIL - ${name}`);
    console.error(err);
    process.exitCode = 1;
  }
}

// 1. Round-trip básico con clave AES-128 (16 bytes)
run('round-trip AES-128 con JSON típico', () => {
  const key = '0123456789abcdef'; // 16 bytes
  const original = JSON.stringify({ sub: 'user-42', role: 'admin', exp: 1999999999 });

  const token = encryptToken(original, key);
  const decrypted = decryptToken(token, key);

  assert.strictEqual(decrypted, original);
  assert.strictEqual(token.length, 16 + Buffer.byteLength(original, 'utf8') * 2);
});

// 2. Round-trip con clave AES-256 (32 bytes)
run('round-trip AES-256', () => {
  const key = '01234567890123456789012345678901'.slice(0, 32); // 32 bytes
  const original = JSON.stringify({ foo: 'bar' });

  const token = encryptToken(original, key);
  const decrypted = decryptToken(token, key);

  assert.strictEqual(decrypted, original);
});

// 3. Cada llamada usa un IV distinto (mismo texto plano -> ciphertext distinto)
run('IV aleatorio produce ciphertexts distintos para el mismo mensaje', () => {
  const key = '0123456789abcdef';
  const original = JSON.stringify({ a: 1 });

  const tokenA = encryptToken(original, key);
  const tokenB = encryptToken(original, key);

  assert.notStrictEqual(tokenA, tokenB);
  assert.strictEqual(decryptToken(tokenA, key), original);
  assert.strictEqual(decryptToken(tokenB, key), original);
});

// 4. Clave incorrecta al descifrar -> el JSON resultante NO coincide (CTR no falla con excepción,
//    solo produce basura: esto ilustra por qué hace falta un MAC para detectar manipulación/errores)
run('clave incorrecta produce texto corrupto en vez de un error', () => {
  const key = '0123456789abcdef';
  const wrongKey = 'fedcba9876543210';
  const original = JSON.stringify({ a: 1 });

  const token = encryptToken(original, key);
  const decrypted = decryptToken(token, wrongKey);

  assert.notStrictEqual(decrypted, original);
});

// 5. getRandomHexString genera la longitud pedida y solo caracteres hex
run('getRandomHexString respeta longitud y alfabeto', () => {
  const hex = getRandomHexString(16);
  assert.strictEqual(hex.length, 16);
  assert.match(hex, /^[0-9a-f]{16}$/);
});

// 6. Longitud de clave inválida debe fallar explícitamente
run('clave de longitud inválida lanza error', () => {
  assert.throws(() => encryptToken('{}', 'clave-muy-corta'));
});
