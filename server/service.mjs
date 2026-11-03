import { assert, AppError, validPlan, codeId, generateCode, nextSubscription, memberStatus, text } from './domain.mjs';

export function createService({ db, auth, isAdmin, clock = Date.now }) {
  const members = db.collection('members'), codes = db.collection('activationCodes');
  const requireAdmin = actor => assert(isAdmin(actor.uid), 403, 'Bạn không có quyền admin.', 'ADMIN_REQUIRED');
  async function me(actor) {
    const ref = members.doc(actor.uid), now = clock();
    const member = await db.runTransaction(async tx => {
      const snap = await tx.get(ref);
      const profile = { uid: actor.uid, email: (actor.email || '').toLowerCase(), name: text(actor.name || actor.email, 120) };
      if (!snap.exists) { const value = { ...profile, createdAt: now, blocked: false, lifetime: false, expiresAt: null }; tx.create(ref, value); return value; }
      const value = snap.data();
      if (value.email !== profile.email || value.name !== profile.name) { tx.update(ref, profile); return { ...value, ...profile }; }
      return value;
    });
    return { uid: actor.uid, email: member.email, name: member.name, emailVerified: actor.email_verified === true, ...memberStatus(member, isAdmin(actor.uid), now), serverTime: now };
  }
  async function requireActive(actor) {
    const account = await me(actor);
    assert(account.emailVerified, 403, 'Xác minh email trước khi sử dụng.', 'EMAIL_UNVERIFIED');
    assert(account.active, 403, account.blocked ? 'Tài khoản đã bị khóa sử dụng.' : 'Bạn cần mã kích hoạt hoặc gói đã hết hạn.', account.blocked ? 'ACCOUNT_BLOCKED' : 'ACTIVATION_REQUIRED');
    return account;
  }
  async function consumeAttempt(uid) {
    const ref = db.collection('rateLimits').doc(uid + '-redeem'), now = clock();
    await db.runTransaction(async tx => {
      const snap = await tx.get(ref), previous = snap.exists ? snap.data() : {};
      const fresh = !previous.startedAt || now - previous.startedAt >= 600000;
      const count = fresh ? 0 : Number(previous.count) || 0;
      assert(count < 20, 429, 'Bạn đã thử quá nhiều mã. Vui lòng đợi 10 phút.', 'RATE_LIMITED');
      tx.set(ref, { startedAt: fresh ? now : previous.startedAt, count: count + 1 });
    });
  }
  async function redeem(actor, input) {
    assert(actor.email_verified === true, 403, 'Xác minh email trước khi nhập mã.', 'EMAIL_UNVERIFIED');
    await consumeAttempt(actor.uid);
    const id = codeId(input.code), codeRef = codes.doc(id), memberRef = members.doc(actor.uid), now = clock();
    await db.runTransaction(async tx => {
      const codeSnap = await tx.get(codeRef), memberSnap = await tx.get(memberRef);
      assert(codeSnap.exists, 400, 'Mã kích hoạt không tồn tại.');
      const code = codeSnap.data(), member = memberSnap.exists ? memberSnap.data() : {};
      assert(code.status === 'unused', 409, code.status === 'redeemed' ? 'Mã này đã được sử dụng và không thể dùng lại.' : 'Mã này đã bị thu hồi.', 'CODE_UNAVAILABLE');
      assert(!member.blocked, 403, 'Tài khoản đã bị khóa sử dụng.', 'ACCOUNT_BLOCKED');
      const subscription = nextSubscription(member, validPlan(code), now);
      tx.set(memberRef, { ...member, uid: actor.uid, email: (actor.email || '').toLowerCase(), name: text(actor.name || actor.email, 120), createdAt: member.createdAt ?? now, blocked: false, ...subscription, lastCodeId: id, updatedAt: now });
      tx.update(codeRef, { status: 'redeemed', usedAt: now, usedByUid: actor.uid, usedByEmail: (actor.email || '').toLowerCase(), subscriptionExpiresAt: subscription.expiresAt, subscriptionLifetime: subscription.lifetime });
      tx.create(db.collection('redemptions').doc(id), { codeId: id, uid: actor.uid, email: (actor.email || '').toLowerCase(), type: code.type, months: code.months, createdAt: now, expiresAt: subscription.expiresAt });
    });
    return me(actor);
  }
  async function createCodes(actor, input) {
    requireAdmin(actor);
    const plan = validPlan(input), count = Number(input.count ?? 1), now = clock();
    assert(Number.isInteger(count) && count >= 1 && count <= 100, 400, 'Mỗi lần tạo từ 1 đến 100 mã.');
    const batch = db.batch(), result = [];
    for (let i = 0; i < count; i++) {
      const code = generateCode(), id = codeId(code), value = { code, ...plan, note: text(input.note), status: 'unused', createdAt: now, createdBy: actor.uid, usedAt: null, usedByUid: null, usedByEmail: null };
      batch.create(codes.doc(id), value); result.push({ id, ...value });
    }
    await batch.commit(); return { codes: result };
  }
  async function revokeCode(actor, input) {
    requireAdmin(actor); assert(/^[a-f0-9]{64}$/.test(input.id || ''), 400, 'ID mã không hợp lệ.');
    const ref = codes.doc(input.id);
    await db.runTransaction(async tx => { const snap = await tx.get(ref); assert(snap.exists, 404, 'Không tìm thấy mã.'); assert(snap.data().status === 'unused', 409, 'Chỉ thu hồi được mã chưa sử dụng.'); tx.update(ref, { status: 'revoked', revokedAt: clock(), revokedBy: actor.uid }); });
    return { ok: true };
  }
  async function summary(actor) {
    requireAdmin(actor);
    const count = async query => (await query.count().get()).data().count;
    const [total, unused, redeemed, revoked, accounts] = await Promise.all([count(codes), count(codes.where('status', '==', 'unused')), count(codes.where('status', '==', 'redeemed')), count(codes.where('status', '==', 'revoked')), count(members)]);
    return { total, unused, redeemed, revoked, accounts, serverTime: clock() };
  }
  async function listCodes(actor, input = {}) {
    requireAdmin(actor);
    const search = text(input.search, 254), status = text(input.status, 20);
    assert(!status || ['unused','redeemed','revoked'].includes(status), 400, 'Bộ lọc không hợp lệ.');
    if (search && !search.includes('@')) {
      const id = codeId(search), snap = await codes.doc(id).get();
      return { items: snap.exists && (!status || snap.data().status === status) ? [{ id, ...snap.data() }] : [], nextCursor: null };
    }
    let query = codes;
    if (status) query = query.where('status', '==', status);
    if (search) query = query.where('usedByEmail', '==', search.toLowerCase());
    query = query.orderBy('createdAt', 'desc');
    if (input.cursor) { assert(/^[a-f0-9]{64}$/.test(input.cursor), 400, 'Cursor không hợp lệ.'); const cursor = await codes.doc(input.cursor).get(); assert(cursor.exists, 400, 'Cursor không tồn tại.'); query = query.startAfter(cursor); }
    const docs = (await query.limit(51).get()).docs, page = docs.slice(0, 50);
    return { items: page.map(doc => ({ id: doc.id, ...doc.data() })), nextCursor: docs.length > 50 ? page.at(-1).id : null };
  }
  async function listMembers(actor, input = {}) {
    requireAdmin(actor);
    if (input.email) { const email = text(input.email,254).toLowerCase(); const user = await auth.getUserByEmail(email).catch(() => null); if (!user) return { items: [], nextCursor: null }; const snap = await members.doc(user.uid).get(); return { items: [{ uid: user.uid, email: user.email, name: user.displayName || '', ...(snap.exists ? snap.data() : {}), ...memberStatus(snap.exists ? snap.data() : {}, isAdmin(user.uid), clock()) }], nextCursor: null }; }
    let query = members.orderBy('createdAt', 'desc');
    if (input.cursor) { assert(/^[A-Za-z0-9_-]{1,128}$/.test(input.cursor), 400, 'Cursor không hợp lệ.'); const cursor = await members.doc(input.cursor).get(); assert(cursor.exists, 400, 'Cursor không tồn tại.'); query = query.startAfter(cursor); }
    const docs = (await query.limit(51).get()).docs, page = docs.slice(0,50);
    return { items: page.map(doc => ({ ...doc.data(), uid: doc.id, ...memberStatus(doc.data(), isAdmin(doc.id), clock()) })), nextCursor: docs.length > 50 ? page.at(-1).id : null };
  }
  async function createMember(actor, input) {
    requireAdmin(actor);
    const email = text(input.email,254).toLowerCase(), password = input.password;
    assert(/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email),400,'Email không hợp lệ.');
    assert(typeof password === 'string' && password.length >= 12 && password.length <= 128,400,'Mật khẩu cấp tài khoản phải từ 12 đến 128 ký tự.');
    let user;
    try { user = await auth.createUser({ email, password, displayName: text(input.name,120) || email, emailVerified: false }); }
    catch (error) { if (error.code === 'auth/email-already-exists') throw new AppError(409,'Email này đã có tài khoản.'); throw error; }
    const member = { uid: user.uid, email, name: user.displayName || email, createdAt: clock(), blocked: false, lifetime: false, expiresAt: null };
    try { await members.doc(user.uid).create(member); } catch(error) { await auth.deleteUser(user.uid).catch(() => {}); throw error; }
    return { user: member };
  }
  async function setBlocked(actor, input) {
    requireAdmin(actor);
    assert(typeof input.uid === 'string' && input.uid.length <= 128 && !input.uid.includes('/'),400,'UID không hợp lệ.');
    assert(typeof input.blocked === 'boolean',400,'Trạng thái không hợp lệ.');
    assert(!isAdmin(input.uid),409,'Không khóa tài khoản admin qua dashboard.');
    const ref=members.doc(input.uid);
    await db.runTransaction(async tx=>{const snap=await tx.get(ref);assert(snap.exists,404,'Không tìm thấy tài khoản.');tx.update(ref,{blocked:input.blocked,updatedAt:clock(),updatedBy:actor.uid});});
    return { ok: true };
  }
  return { me, requireActive, redeem, createCodes, revokeCode, summary, listCodes, listMembers, createMember, setBlocked };
}
