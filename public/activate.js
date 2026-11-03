import {
  auth,
  currentUser,
  api,
  safeNext,
  authError,
  sendEmailVerification,
  signOut,
  reload
} from './firebase-client.js';

const ADMIN_FACEBOOK = 'https://www.facebook.com/cossiii.py';

const $ = id => document.getElementById(id);

const next = safeNext(
  new URLSearchParams(location.search).get('next')
);

const message = value => {
  $('activation-message').textContent = value;
};

function busy(value) {
  document.querySelectorAll('button').forEach(button => {
    button.disabled = value;
  });
}

function showSubscriptionInfo(text) {
  const info = $('subscription-info');
  const position = text.indexOf('admin');

  if (position === -1) {
    info.textContent = text;
    return;
  }

  const link = document.createElement('a');

  link.href = ADMIN_FACEBOOK;
  link.target = '_blank';
  link.rel = 'noopener noreferrer';
  link.textContent = 'admin';
  link.style.color = '#107c41';
  link.style.textDecoration = 'underline';

  info.replaceChildren(
    document.createTextNode(text.slice(0, position)),
    link,
    document.createTextNode(text.slice(position + 5))
  );
}

async function refresh() {
  const user = await currentUser();

  if (!user) {
    location.replace(
      '/login.html?next=' + encodeURIComponent(next)
    );
    return;
  }

  const me = await api('me');

  $('account-email').textContent = me.email;

  $('verify-area').hidden = me.emailVerified;

  $('code-form').hidden =
    !me.emailVerified ||
    me.blocked ||
    me.lifetime ||
    me.admin;

  $('continue-link').hidden =
    !me.active || !me.emailVerified;

  $('continue-link').href = next;

  $('admin-link').hidden =
    !me.admin || !me.emailVerified;

  let subscriptionText;

  if (me.blocked) {
    subscriptionText =
      'Tài khoản đã bị khóa sử dụng. Liên hệ admin.';
  } else if (!me.emailVerified) {
    subscriptionText = 'Email chưa được xác minh.';
  } else if (me.admin) {
    subscriptionText = 'Tài khoản quản trị.';
  } else if (me.lifetime) {
    subscriptionText = 'Bạn có quyền sử dụng vĩnh viễn.';
  } else if (me.active) {
    subscriptionText =
      'Đang hoạt động đến ' +
      new Date(me.expiresAt).toLocaleString('vi-VN');
  } else if (me.expiresAt) {
    subscriptionText =
      'Gói đã hết hạn. Nhập mã mới để gia hạn.';
  } else {
    subscriptionText =
      'Bạn chưa có gói sử dụng. Nhập mã do admin cấp.';
  }

  showSubscriptionInfo(subscriptionText);

  return me;
}

$('code-form').onsubmit = async event => {
  event.preventDefault();

  busy(true);
  message('Đang kích hoạt…');

  try {
    await api('redeem', {
      code: $('activation-code').value
    });

    message('Kích hoạt thành công.');
    location.replace(next);
  } catch (error) {
    message(authError(error));
  } finally {
    busy(false);
  }
};

$('send-verify').onclick = async () => {
  busy(true);

  try {
    await sendEmailVerification(auth.currentUser);

    message(
      'Đã gửi email xác minh. Kiểm tra cả thư mục Spam.'
    );
  } catch (error) {
    message(authError(error));
  } finally {
    busy(false);
  }
};

$('check-verify').onclick = async () => {
  busy(true);

  try {
    await reload(auth.currentUser);
    await auth.currentUser.getIdToken(true);

    const me = await refresh();

    if (!me) return;

    message(
      me.emailVerified
        ? 'Email đã xác minh. Bạn có thể nhập mã.'
        : 'Email vẫn chưa xác minh. Mở liên kết trong email rồi thử lại.'
    );
  } catch (error) {
    message(authError(error));
  } finally {
    busy(false);
  }
};

$('logout').onclick = async () => {
  await signOut(auth);
  location.replace('/login.html');
};

refresh().catch(error => {
  message(authError(error));
});