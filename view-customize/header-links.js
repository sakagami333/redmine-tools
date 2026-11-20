/* ==========================================================
   ヘッダーに独自のリンクを追加する(Redmine 7.0対応版)

   View customize の設定:
     パスのパターン       :(空欄)
     プロジェクトのパターン:(空欄)
     挿入位置             :全ページのヘッダ
     種別                 :JavaScript

   LINKS の query_id は、自分の環境のカスタムクエリの番号に書き換えること
   (カスタムクエリを開いたときのURLの query_id=〇〇)。
   ========================================================== */
(function () {
  // 追加するリンク(上から順に右へ並ぶ)
  const LINKS = [
    { text: 'いつかやる',         href: '/issues?query_id=50' },
    { text: '自分の未完了',       href: '/issues?query_id=49' },
    { text: '進行中のタスク',     href: '/issues?query_id=53' },
    { text: '進行中のストーリー', href: '/issues?query_id=52' },
  ];

  function addLinks() {
    if (document.querySelector('#top-menu .vc-custom-link')) return; // 二重追加を防ぐ

    // 「ヘルプ」(なければ「ホーム」)のリンクを手がかりに、ヘッダーのメニューを探す
    const base = document.querySelector('#top-menu a.help, #top-menu a.home');
    const baseItem = base ? base.closest('li') : null;
    const list = baseItem ? baseItem.parentNode : document.querySelector('#top-menu ul');
    if (!list) {
      console.log('[ヘッダーリンク] メニューが見つかりません');
      return;
    }

    LINKS.forEach(({ text, href }) => {
      // 既存の項目を複製して、見た目をそろえる
      const li = baseItem ? baseItem.cloneNode(true) : document.createElement('li');
      li.className = '';
      let a = li.querySelector('a');
      if (!a) {
        a = document.createElement('a');
        li.appendChild(a);
      }
      a.className = 'vc-custom-link';
      a.removeAttribute('target');
      a.href = href;
      a.textContent = text;
      list.appendChild(li);
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', addLinks);
  } else {
    addLinks();
  }
})();
