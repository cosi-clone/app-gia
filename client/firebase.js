import { initializeApp } from 'firebase/app';
import { getAuth, GoogleAuthProvider, onAuthStateChanged, signInWithPopup, signInWithEmailAndPassword, createUserWithEmailAndPassword, sendEmailVerification, sendPasswordResetEmail, signOut, reload, connectAuthEmulator } from 'firebase/auth';

export let auth, settings;
export const firebaseReady=(async()=>{
  const response=await fetch('/api/service?route=config',{cache:'no-store'});const data=await response.json();if(!response.ok)throw Error(data.error||'Chưa cấu hình Firebase.');
  settings=data;auth=getAuth(initializeApp(data.firebase));
  if(data.emulator)connectAuthEmulator(auth,'http://127.0.0.1:9099',{disableWarnings:true});
  return auth;
})();
// Avoid unhandled rejection while the page attaches its own error handler.
firebaseReady.catch(()=>{});
export async function currentUser(){await firebaseReady;return new Promise(resolve=>{let off=()=>{};off=onAuthStateChanged(auth,user=>{off();resolve(user);});});}
export async function api(route,body,params={}){
  const user=await currentUser();if(!user)throw Object.assign(Error('Vui lòng đăng nhập.'),{status:401,code:'LOGIN_REQUIRED'});
  const query=new URLSearchParams({route,...params});
  const response=await fetch('/api/service?'+query,{method:body===undefined?'GET':'POST',headers:{Authorization:'Bearer '+await user.getIdToken(),...(body===undefined?{}:{'Content-Type':'application/json'})},body:body===undefined?undefined:JSON.stringify(body),cache:'no-store'});
  const data=await response.json();if(!response.ok)throw Object.assign(Error(data.error||'Yêu cầu thất bại.'),{status:response.status,code:data.code});return data;
}
export function safeNext(value,fallback='/index.html'){
  try{const url=new URL(value||fallback,location.origin);if(url.origin!==location.origin||!['/','/index.html','/reader.html','/admin.html'].includes(url.pathname))return fallback;return url.pathname+url.search+url.hash;}catch{return fallback;}
}
export function authError(error){const errors={'auth/invalid-credential':'Email hoặc mật khẩu không đúng.','auth/email-already-in-use':'Email đã được đăng ký.','auth/weak-password':'Mật khẩu chưa đủ mạnh.','auth/invalid-email':'Email không hợp lệ.','auth/too-many-requests':'Bạn thử quá nhiều lần. Vui lòng thử lại sau.','auth/popup-closed-by-user':'Cửa sổ đăng nhập đã đóng.','auth/operation-not-allowed':'Nhà cung cấp đăng nhập chưa được bật trong Firebase Console.','auth/unauthorized-domain':'Thêm domain hiện tại vào Firebase Authentication → Authorized domains.','auth/network-request-failed':'Không kết nối được Firebase. Kiểm tra mạng.','auth/user-disabled':'Tài khoản đã bị vô hiệu hóa.'};return errors[error.code]||error.message||'Có lỗi xảy ra.';}
export {GoogleAuthProvider,onAuthStateChanged,signInWithPopup,signInWithEmailAndPassword,createUserWithEmailAndPassword,sendEmailVerification,sendPasswordResetEmail,signOut,reload};
