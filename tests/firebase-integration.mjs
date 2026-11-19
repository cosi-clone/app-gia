import test from 'node:test';
import assert from 'node:assert/strict';
import { initializeApp, deleteApp } from 'firebase-admin/app';
import { getFirestore } from 'firebase-admin/firestore';
import { createService } from '../server/service.mjs';

const enabled=Boolean(process.env.FIRESTORE_EMULATOR_HOST);
test('real Firestore emulator: concurrent single-use redemption, persisted owner and no partial grant',{skip:!enabled},async()=>{
  const projectId='demo-excel-reader',app=initializeApp({projectId},'integration-'+Date.now()),db=getFirestore(app);
  const suffix=Date.now().toString(36),admin={uid:'integration-admin-'+suffix,email:'admin@example.com',email_verified:true},alice={uid:'integration-alice-'+suffix,email:'alice@example.com',email_verified:true},bob={uid:'integration-bob-'+suffix,email:'bob@example.com',email_verified:true};
  const service=createService({db,auth:{},isAdmin:uid=>uid===admin.uid});
  try{const code=(await service.createCodes(admin,{type:'months',months:3})).codes[0];const attempts=await Promise.allSettled([service.redeem(alice,{code:code.code}),service.redeem(bob,{code:code.code})]);assert.equal(attempts.filter(r=>r.status==='fulfilled').length,1);assert.equal(attempts.filter(r=>r.status==='rejected'&&r.reason.status===409).length,1);const stored=(await db.collection('activationCodes').doc(code.id).get()).data();assert.equal(stored.status,'redeemed');assert.ok([alice.uid,bob.uid].includes(stored.usedByUid));const audit=await db.collection('redemptions').doc(code.id).get();assert.equal(audit.data().uid,stored.usedByUid);await assert.rejects(()=>service.redeem(stored.usedByUid===alice.uid?alice:bob,{code:code.code}),{status:409});}finally{await db.terminate();await deleteApp(app);}
});
