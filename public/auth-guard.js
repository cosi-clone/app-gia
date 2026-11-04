import {
  auth,
  currentUser,
  api,
  signOut,
  onAuthStateChanged,
  safeNext
} from './firebase-client.js';

let uid;
let account;
let unsubscribe;
let accessPromise;

const next = () => safeNext(
  location.pathname + location.search + location.hash
);

function navigate(path) {
  document.body.classList.add('access-loading');

  location.replace(
    path + '?next=' + encodeURIComponent(next())
  );
}

function showError(error) {
  document.body.classList.add('access-loading');

  let box = document.getElementById('access-error');

  if (!box) {
    box = document.createElement('div');
    box.id = 'access-error';
    box.className = 'access-error';
    document.body.append(box);
  }

  const p = document.createElement('p');
  p.textContent = error.message;

  const retry = document.createElement('a');
  retry.href = location.href;
  retry.textContent = 'Thử lại';

  const login = document.createElement('a');
  login.href = '/login.html';
  login.textContent = 'Đăng nhập';

  box.replaceChildren(p, retry, login);
}

async function requireActive() {
  try {
    const user = await currentUser();

    if (!user) {
      throw Object.assign(
        Error('Vui lòng đăng nhập.'),
        { status: 401 }
      );
    }

    if (uid && uid !== user.uid) {
      throw Object.assign(
        Error('Tài khoản đã thay đổi.'),
        { status: 401 }
      );
    }

    uid = user.uid;

    // Chỉ kiểm tra API một lần trong lần tải trang này.
    // Reload hoặc mở tab mới sẽ kiểm tra lại.
    if (!accessPromise) {
      accessPromise = api('access').catch(error => {
        accessPromise = null;
        throw error;
      });
    }

    account = await accessPromise;

    return account;
  } catch (error) {
    if (error.status === 401) {
      navigate('/login.html');
    } else if (error.status === 403) {
      navigate('/activate.html');
    } else {
      showError(error);
    }

    throw error;
  }
}

async function requireAdmin() {
  const user = await currentUser();

  if (!user) {
    navigate('/login.html');
    throw Error('Vui lòng đăng nhập.');
  }

  uid = user.uid;
  account = await api('me');

  if (!account.admin) {
    location.replace('/index.html');
    throw Error('Không có quyền admin.');
  }

  if (!account.emailVerified) {
    navigate('/activate.html');
    throw Error('Xác minh email trước.');
  }

  return account;
}

function startMonitor() {
  // Chỉ theo dõi đăng xuất hoặc đổi tài khoản.
  // Không kiểm tra mỗi phút, không kiểm tra khi quay lại tab.
  if (unsubscribe) {
    unsubscribe();
  }

  unsubscribe = onAuthStateChanged(auth, user => {
    if (!user || user.uid !== uid) {
      accessPromise = null;
      account = null;
      navigate('/login.html');
    }
  });
}

function accountBar(target) {
  const bar = document.createElement('div');
  bar.className = 'account-bar';

  const email = document.createElement('span');
  email.textContent = account.email;

  const plan = document.createElement('span');

  plan.textContent = account.admin
    ? 'Admin'
    : account.lifetime
      ? 'Vĩnh viễn'
      : 'Hết hạn: ' +
        new Date(account.expiresAt)
          .toLocaleDateString('vi-VN');

  const renew = document.createElement('a');
  renew.href = '/activate.html';
  renew.textContent = 'Tài khoản / Gia hạn';

  bar.append(email, plan, renew);

  if (account.admin) {
    const admin = document.createElement('a');
    admin.href = '/admin.html';
    admin.textContent = 'Quản trị';
    bar.append(admin);
  }

  const out = document.createElement('button');
  out.textContent = 'Đăng xuất';

  out.onclick = async () => {
    accessPromise = null;
    await signOut(auth);
    location.replace('/login.html');
  };

  bar.append(out);
  target.prepend(bar);
}

window.Access = {
  requireActive,
  requireAdmin,

  get uid() {
    return uid;
  },

  get account() {
    return account;
  }
};

export {
  requireActive,
  requireAdmin,
  startMonitor,
  accountBar,
  showError
};