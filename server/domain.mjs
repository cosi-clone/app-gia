import { createHash, randomBytes } from 'node:crypto';

export class AppError extends Error {
  constructor(status, message, code = 'REQUEST_FAILED') { super(message); this.status = status; this.code = code; }
}
export function assert(condition, status, message, code) { if (!condition) throw new AppError(status, message, code); }
export function addCalendarMonths(timestamp, months) {
  const offset = 7 * 60 * 60 * 1000;
  const date = new Date(timestamp + offset), day = date.getUTCDate();
  date.setUTCDate(1); date.setUTCMonth(date.getUTCMonth() + months);
  const lastDay = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + 1, 0)).getUTCDate();
  date.setUTCDate(Math.min(day, lastDay)); return date.getTime() - offset;
}
export function validPlan(input) {
  assert(['months', 'lifetime'].includes(input.type), 400, 'Chọn gói theo tháng hoặc vĩnh viễn.');
  const months = Number(input.months);
  assert(input.type === 'lifetime' || Number.isInteger(months) && months >= 1 && months <= 1200, 400, 'Số tháng phải là số nguyên từ 1 đến 1200.');
  return { type: input.type, months: input.type === 'months' ? months : null };
}
export function normalizeCode(value) {
  assert(typeof value === 'string', 400, 'Nhập mã kích hoạt.');
  const code = value.trim().toUpperCase().replace(/[\s-]/g, '');
  assert(/^EX[0-9A-F]{36}$/.test(code), 400, 'Mã kích hoạt không đúng định dạng.');
  return code;
}
export function codeId(value) { return createHash('sha256').update(normalizeCode(value)).digest('hex'); }
export function generateCode() { return 'EX-' + randomBytes(18).toString('hex').toUpperCase().match(/.{6}/g).join('-'); }
export function nextSubscription(current, plan, now) {
  if (plan.type === 'lifetime') return { lifetime: true, expiresAt: null, activatedAt: current.activatedAt ?? now };
  assert(!current.lifetime, 409, 'Tài khoản đã có quyền vĩnh viễn; mã chưa bị sử dụng.');
  return { lifetime: false, expiresAt: addCalendarMonths(Math.max(now, Number(current.expiresAt) || 0), plan.months), activatedAt: current.activatedAt ?? now };
}
export function memberStatus(member, admin, now = Date.now()) {
  const blocked = member.blocked === true;
  return { admin, blocked, active: !blocked && (admin || member.lifetime === true || Number(member.expiresAt) > now), lifetime: member.lifetime === true, expiresAt: member.lifetime ? null : member.expiresAt ?? null };
}
export function text(value, max = 200) { return typeof value === 'string' ? value.trim().slice(0, max) : ''; }
