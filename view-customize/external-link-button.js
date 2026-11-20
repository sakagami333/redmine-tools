/* ==========================================================
   編集ツールバー:「Wikiページへのリンク」ボタンを「外部リンク」ボタンに置き換える
   (Redmine 7.0 / テキスト書式 CommonMark Markdown 用)

   View customize の設定:
     パスのパターン       :(空欄)
     プロジェクトのパターン:(空欄)
     挿入位置             :全ページのヘッダ
     種別                 :JavaScript

   ボタンを押したときの動き:
     - 文字を選んで押す        → URL を聞いて [選んだ文字](URL) にする
     - URL を選んで押す        → [URL](URL) にして、表示する文字の部分を選択状態にする
     - 何も選ばずに押す        → URL を聞いて [URL](URL) を入れ、表示する文字の部分を選択状態にする
   ========================================================== */
(function () {
  'use strict';

  var TITLE = '外部リンク';
  var URL_PATTERN = /^(https?:\/\/|mailto:)\S+$/i;

  // ボタンのアイコン(「外部リンク」の図柄。Redmine 7.0 の他のアイコンと同じ線の太さ)
  var ICON =
    'data:image/svg+xml;utf8,' +
    encodeURIComponent(
      '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" ' +
        'fill="none" stroke="#333" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">' +
        '<path d="M12 6h-6a2 2 0 0 0 -2 2v10a2 2 0 0 0 2 2h10a2 2 0 0 0 2 -2v-6"/>' +
        '<path d="M11 13l9 -9"/><path d="M15 4h5v5"/></svg>'
    );

  function addStyle() {
    if (document.getElementById('vc-external-link-style')) return;
    var style = document.createElement('style');
    style.id = 'vc-external-link-style';
    style.textContent = '.jstElements button.jstb_link { background-image: url("' + ICON + '") !important; }';
    document.head.appendChild(style);
  }

  // 選択範囲を text に置き換え、置き換え後の [selStart, selEnd] を選択する(元に戻す操作も効くように)
  function replaceSelection(textarea, text, selStart, selEnd) {
    var start = textarea.selectionStart;
    textarea.focus();
    var ok = false;
    try {
      ok = document.execCommand('insertText', false, text);
    } catch (e) {
      ok = false;
    }
    if (!ok) {
      textarea.setRangeText(text, start, textarea.selectionEnd, 'end');
      textarea.dispatchEvent(new Event('input', { bubbles: true }));
    }
    textarea.setSelectionRange(start + selStart, start + selEnd);
  }

  function insertExternalLink(textarea) {
    if (!textarea) return;
    var selected = textarea.value.substring(textarea.selectionStart, textarea.selectionEnd);
    var trailing = (selected.match(/\s+$/) || [''])[0]; // 選択の末尾の空白は、リンクの外に残す
    var label = selected.slice(0, selected.length - trailing.length).trim();

    // URL を選んでいる場合:そのまま使い、表示する文字を選択状態にする
    if (URL_PATTERN.test(label)) {
      replaceSelection(textarea, '[' + label + '](' + label + ')' + trailing, 1, 1 + label.length);
      return;
    }

    var url = window.prompt('リンク先のURLを入力してください', 'https://');
    if (url === null) return; // キャンセル
    url = url.trim();
    if (!url || url === 'https://') return;

    if (label) {
      // 文字を選んでいる場合:その文字をリンクにする(カーソルはリンクの後ろへ)
      var text = '[' + label + '](' + url + ')';
      replaceSelection(textarea, text + trailing, text.length, text.length);
    } else {
      // 何も選んでいない場合:URL を表示する文字にし、そこを選択状態にする(上書きで変えられる)
      replaceSelection(textarea, '[' + url + '](' + url + ')', 1, 1 + url.length);
    }
  }

  // ツールバーのボタンを置き換える(後から表示される編集欄にも対応)
  function upgradeButtons(root) {
    var buttons = (root.querySelectorAll ? root : document).querySelectorAll('.jstElements button.jstb_link');
    buttons.forEach(function (button) {
      if (button.dataset.vcExternalLink) return;
      button.dataset.vcExternalLink = '1';
      button.title = TITLE;
      var label = button.querySelector('span');
      if (label) label.textContent = TITLE;
      button.onclick = function () {
        var block = button.closest('.jstBlock');
        insertExternalLink(block && block.querySelector('textarea'));
        return false;
      };
    });
  }

  function start() {
    addStyle();
    upgradeButtons(document);
    new MutationObserver(function (mutations) {
      for (var i = 0; i < mutations.length; i++) {
        if (mutations[i].addedNodes.length) {
          upgradeButtons(document);
          return;
        }
      }
    }).observe(document.body, { childList: true, subtree: true });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', start);
  } else {
    start();
  }
})();
