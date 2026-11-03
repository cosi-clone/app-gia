import 'dotenv/config';
import { backend } from '../server/firebase.mjs';
const email=process.env.BOOTSTRAP_ADMIN_EMAIL,password=process.env.BOOTSTRAP_ADMIN_PASSWORD;
if(!email){console.error('Đặt BOOTSTRAP_ADMIN_EMAIL trong .env rồi chạy lại.');process.exit(1);}
try{
  const {auth}=backend();let user;
  try{user=await auth.getUserByEmail(email);}catch(error){if(error.code!=='auth/user-not-found')throw error;if(!password||password.length<12)throw Error('Tài khoản mới cần BOOTSTRAP_ADMIN_PASSWORD từ 12 ký tự.');user=await auth.createUser({email,password,emailVerified:true,displayName:'Administrator'});}
  if(!user.emailVerified)user=await auth.updateUser(user.uid,{emailVerified:true});
  console.log(`Admin UID: ${user.uid}\nThêm UID này vào ADMIN_UIDS trong .env và Vercel, rồi khởi động/deploy lại.\nXóa BOOTSTRAP_ADMIN_PASSWORD sau khi tạo xong.`);
}catch(error){console.error(error.message);process.exit(1);}
