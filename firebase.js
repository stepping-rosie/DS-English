// ═══════════════════════════════════════════════════════
//  DS English — Firebase 공통 데이터 레이어
//  firebase.js  (admin / teacher / parent / student 공유)
// ═══════════════════════════════════════════════════════

// ▼▼▼ Firebase 콘솔에서 복사한 설정값을 여기에 붙여넣으세요 ▼▼▼
const FIREBASE_CONFIG = {
  apiKey: "AIzaSyCVL8HSLgsNUa37ih6uzPqCj3yjQ7JxsHI",
  authDomain: "ds-english-cc3dd.firebaseapp.com",
  projectId: "ds-english-cc3dd",
  storageBucket: "ds-english-cc3dd.firebasestorage.app",
  messagingSenderId: "146017984739",
  appId: "1:146017984739:web:6ad57f1c42589e299f65d6",
  measurementId: "G-Z7SE1S35JK"
};
// ▲▲▲ 설정값 끝 ▲▲▲

// Firestore에 저장할 키 목록 (localStorage 키와 동일)
const DS_KEYS = [
  'students','classes','teachers','fees','scores',
  'messages','teacherMessages','studentMessages',
  'att','teacherAtt','studentBadges','announcements',
  'homework','examTypes','teacherChat','admin_pw'
];

// 기본값이 {} 인 키 (나머지는 [])
const OBJ_KEYS = new Set(['att','teacherAtt']);

// 전역 캐시 — 모든 데이터를 메모리에 보관 (읽기 동기화용)
window.CACHE = {};
window._fbDb  = null;

// ── 읽기 함수 (동기, 캐시 기반) ────────────────────────
function _l(k) {
  const v = window.CACHE[k];
  return Array.isArray(v) ? v : [];
}
function _lo(k) {
  const v = window.CACHE[k];
  return (v && typeof v === 'object' && !Array.isArray(v)) ? v : {};
}

// ── 쓰기 함수 (캐시 즉시 반영 + Firestore 비동기 저장) ─
function fbWrite(k, v) {
  window.CACHE[k] = v;
  if (window._fbDb) {
    window._fbDb.collection('ds_data').doc(k)
      .set({ value: v })
      .catch(e => console.error('[Firebase] 쓰기 오류:', k, e));
  }
}
function _s(k, v)  { fbWrite(k, v); }
function _so(k, v) { fbWrite(k, v); }

// ── Firebase 초기화 + 전체 데이터 로딩 ─────────────────
async function fbInit() {
  try {
    if (!firebase.apps.length) {
      firebase.initializeApp(FIREBASE_CONFIG);
    }
    const db = firebase.firestore();
    window._fbDb = db;

    // 모든 키 병렬 로딩
    await Promise.all(DS_KEYS.map(async k => {
      const snap = await db.collection('ds_data').doc(k).get();
      window.CACHE[k] = snap.exists
        ? snap.data().value
        : (OBJ_KEYS.has(k) ? {} : []);
    }));

    // 실시간 변경 감지 — 다른 컴퓨터의 업데이트를 자동 반영
    DS_KEYS.forEach(k => {
      db.collection('ds_data').doc(k).onSnapshot(snap => {
        window.CACHE[k] = snap.exists
          ? snap.data().value
          : (OBJ_KEYS.has(k) ? {} : []);
        // 각 페이지에서 정의한 갱신 콜백 호출
        if (typeof window._onFbUpdate === 'function') {
          window._onFbUpdate(k);
        }
      });
    });

    return true;
  } catch (e) {
    console.error('[Firebase] 초기화 오류:', e);
    return false;
  }
}

// ── localStorage → Firebase 데이터 이전 (최초 1회) ─────
async function migrateFromLocalStorage() {
  if (!window._fbDb) return 0;
  const db = window._fbDb;
  let count = 0;
  for (const k of DS_KEYS) {
    if (k === 'admin_pw') continue;
    const raw = localStorage.getItem('ds_' + k);
    if (!raw) continue;
    try {
      const val = JSON.parse(raw);
      if (Array.isArray(val) && !val.length) continue;
      if (!Array.isArray(val) && typeof val === 'object' && !Object.keys(val).length) continue;
      await db.collection('ds_data').doc(k).set({ value: val });
      window.CACHE[k] = val;
      count++;
    } catch (e) {
      console.error('[이전 오류]', k, e);
    }
  }
  return count;
}
