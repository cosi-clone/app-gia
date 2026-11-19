import test from 'node:test';
import assert from 'node:assert/strict';
import { webConfig, adminChecker } from '../server/firebase.mjs';
const config=JSON.stringify({apiKey:'public',authDomain:'demo.firebaseapp.com',projectId:'demo',appId:'demo-app',private_key:'MUST_NOT_LEAK'});
test('public config never exposes admin credentials or arbitrary fields',()=>{const result=webConfig({FIREBASE_WEB_CONFIG:config,FIREBASE_SERVICE_ACCOUNT:'SECRET'});assert.equal(result.firebase.apiKey,'public');assert.equal(result.firebase.private_key,undefined);assert.ok(!JSON.stringify(result).includes('SECRET'));});
test('emulators cannot be enabled on Vercel or production',()=>{for(const extra of [{NODE_ENV:'production'},{VERCEL:'1'}])assert.throws(()=>webConfig({FIREBASE_WEB_CONFIG:config,USE_FIREBASE_EMULATORS:'true',...extra}),{status:503});});
test('admin role only comes from UID allowlist',()=>{const allowed=adminChecker({ADMIN_UIDS:' uid-a,uid-b '});assert.equal(allowed('uid-a'),true);assert.equal(allowed('uid-b'),true);assert.equal(allowed('admin@example.com'),false);assert.equal(adminChecker({})('uid-a'),false);});
