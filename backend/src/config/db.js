import mysql from 'mysql2/promise';
import dotenv from 'dotenv';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

dotenv.config();

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// SSL config: always use SSL for non-localhost connections (Railway, Aiven, etc.)
function getSSLConfig() {
  const host = process.env.DB_HOST || 'localhost';
  
  // Non-localhost connections require SSL
  if (host !== 'localhost' && host !== '127.0.0.1') {
    return { rejectUnauthorized: false };
  }
  return undefined;
}

let _pool = null;

function getPool() {
  if (!_pool) {
    const ssl = getSSLConfig();
    console.log('Creating pool - SSL:', ssl ? 'enabled' : 'disabled', 'Host:', process.env.DB_HOST);
    
    _pool = mysql.createPool({
      host: process.env.DB_HOST || 'localhost',
      port: Number(process.env.DB_PORT) || 3306,
      user: process.env.DB_USER || 'root',
      password: process.env.DB_PASSWORD || '',
      database: process.env.DB_NAME || 'db_absensi',
      waitForConnections: true,
      connectionLimit: 10,
      queueLimit: 0,
      dateStrings: true,
      ssl: ssl,
    });
  }
  return _pool;
}

// Proxy agar controller tetap bisa panggil pool.query() langsung
const pool = new Proxy({}, {
  get(_, prop) {
    const p = getPool();
    const val = p[prop];
    if (typeof val === 'function') {
      return val.bind(p);
    }
    return val;
  }
});

export default pool;