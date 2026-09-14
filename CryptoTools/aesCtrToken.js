'use strict';

/**
 * Port 1:1 del código Java original (AES/CTR/NoPadding), bugs incluidos,
 * para poder interoperar con tokens generados por ese código.
 *
 * Bugs conservados a propósito (ver revisión de seguridad):
 *  - El IV es la cadena hex en ASCII, NO los bytes que esa cadena hex representa.
 *  - No hay autenticación/integridad (CTR es maleable).
 *  - La clave se usa tal cual como bytes (debe medir 16/24/32 bytes).
 *
 * NOTA: el método original `getRandomHexString(randomLength)` no fue provisto,
 * así que aquí hay una implementación genérica equivalente en comportamiento
 * (genera una cadena hex en minúsculas de la longitud pedida usando el CSPRNG
 * del sistema). Si el original usaba otra fuente de aleatoriedad (p. ej.
 * java.util.Random sin semilla segura), reemplázala aquí para igualar bit a bit.
 */

const crypto = require('crypto');

// AES-CTR exige un IV de exactamente 16 bytes (tamaño de bloque de AES).
// Como el IV son los bytes ASCII de randomHex (bug conservado), randomHex
// debe tener exactamente 16 caracteres.
const RANDOM_HEX_LENGTH = 16;

function getRandomHexString(length) {
  const bytesNeeded = Math.ceil(length / 2);
  return crypto.randomBytes(bytesNeeded).toString('hex').slice(0, length);
}

function algorithmForKeyLength(keyLengthBytes) {
  switch (keyLengthBytes) {
    case 16: return 'aes-128-ctr';
    case 24: return 'aes-192-ctr';
    case 32: return 'aes-256-ctr';
    default:
      throw new Error(
        `Longitud de clave AES inválida: ${keyLengthBytes} bytes (se esperaba 16, 24 o 32)`
      );
  }
}

/**
 * @param {string} jsonString  Texto plano a cifrar (p. ej. un JSON serializado).
 * @param {string} secretKey   Clave AES como string (sus bytes UTF-8 se usan tal cual).
 * @returns {string} randomHex (16 chars) + ciphertext en hex, igual que el Java original.
 */
function encryptToken(jsonString, secretKey) {
  const randomHex = getRandomHexString(RANDOM_HEX_LENGTH);
  const key = Buffer.from(secretKey, 'utf8');
  const iv = Buffer.from(randomHex, 'utf8'); // bug conservado: ASCII del hex, no hex decodificado
  const algorithm = algorithmForKeyLength(key.length);

  const cipher = crypto.createCipheriv(algorithm, key, iv);
  const encrypted = Buffer.concat([cipher.update(jsonString, 'utf8'), cipher.final()]);

  return randomHex + encrypted.toString('hex');
}

/**
 * @param {string} token      Salida de encryptToken: randomHex (16 chars) + ciphertext en hex.
 * @param {string} secretKey  Misma clave usada para cifrar.
 * @returns {string} jsonString original.
 */
function decryptToken(token, secretKey) {
  if (token.length <= RANDOM_HEX_LENGTH) {
    throw new Error('Token demasiado corto: falta el IV o el ciphertext');
  }

  const randomHex = token.slice(0, RANDOM_HEX_LENGTH);
  const cipherHex = token.slice(RANDOM_HEX_LENGTH);

  const key = Buffer.from(secretKey, 'utf8');
  const iv = Buffer.from(randomHex, 'utf8');
  const algorithm = algorithmForKeyLength(key.length);

  const decipher = crypto.createDecipheriv(algorithm, key, iv);
  const decrypted = Buffer.concat([
    decipher.update(Buffer.from(cipherHex, 'hex')),
    decipher.final(),
  ]);

  return decrypted.toString('utf8');
}

module.exports = { encryptToken, decryptToken, getRandomHexString };
