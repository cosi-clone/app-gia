import { getApps, initializeApp, cert } from 'firebase-admin/app';
import { getAuth } from 'firebase-admin/auth';
import { getFirestore } from 'firebase-admin/firestore';
import { AppError } from './domain.mjs';

export function webConfig(env = process.env) {
  let config;
  try { config = JSON.parse(env.FIREBASE_WEB_CONFIG || '{}'); } catch { throw new AppError(503, 'FIREBASE_WEB_CONFIG chưa đúng định dạng JSON.'); }
  if (!config.apiKey || !config.authDomain || !config.projectId || !config.appId) throw new AppError(503, 'Chưa cấu hình Firebase. Xem SETUP_FIREBASE.md trong project.');
  const safe = {};
  for (const key of ['apiKey','authDomain','projectId','appId','messagingSenderId','storageBucket','measurementId']) if (config[key]) safe[key] = String(config[key]);
  const emulator = env.USE_FIREBASE_EMULATORS === 'true';
  if (emulator && (env.NODE_ENV === 'production' || env.VERCEL)) throw new AppError(503, 'Không được bật Firebase Emulator trên bản deploy.');
  return { firebase: safe, googleEnabled: env.ENABLE_GOOGLE_LOGIN === 'true', emulator };
}
export function backend(env = process.env) {
  const { firebase, emulator } = webConfig(env);
  if (!getApps().length) {
    if (emulator) {
      process.env.FIREBASE_AUTH_EMULATOR_HOST = '127.0.0.1:9099';
      process.env.FIRESTORE_EMULATOR_HOST = '127.0.0.1:8080';
      initializeApp({ projectId: firebase.projectId });
    } else {
      let service;
      try { service = JSON.parse(env.FIREBASE_SERVICE_ACCOUNT || '{}'); } catch { throw new AppError(503, 'FIREBASE_SERVICE_ACCOUNT chưa đúng định dạng JSON.'); }
      if (!service.private_key || !service.client_email || service.project_id !== firebase.projectId) throw new AppError(503, 'Service account chưa được cấu hình hoặc không cùng Firebase project.');
      service.private_key = service.private_key.replace(/\\n/g, '\n');
      initializeApp({ credential: cert(service), projectId: firebase.projectId });
    }
  }
  return { auth: getAuth(), db: getFirestore() };
}
export function adminChecker(env = process.env) {
  const ids = new Set((env.ADMIN_UIDS || '').split(',').map(v => v.trim()).filter(Boolean));
  return uid => ids.has(uid);
}
