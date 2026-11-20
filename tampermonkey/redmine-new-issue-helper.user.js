// ==UserScript==
// @name         Redmine 新しいチケットの作成補助
// @description  新規チケットの担当者を自分にし、子チケット作成時はトラッカーを1階層下に設定して、親のタイトルと説明の雛形を入力する
// @version      3
// @grant        none
// @match        http://localhost:3000/projects/*/issues/new*
// @match        http://localhost:3001/projects/*/issues/new*
// @run-at       document-idle
// ==/UserScript==

(async function () {
  'use strict';

  // ===== 設定 =====
  // 親チケットのトラッカー名 → 子チケットに設定するトラッカー名
  // (ここにない親トラッカーの場合、トラッカーは変更しない)
  const CHILD_TRACKER = {
    'ストーリー': '作業区分',
    '作業区分': 'タスク',
  };

  // 子チケットの説明欄に入れる雛形
  const DESCRIPTION_TEMPLATE = '# このタスクのゴール\n\n\n\n# 対応内容まとめ\n\n\n\n';

  const LOG = '[チケット作成補助]';
  const byId = (id, doc = document) => doc.getElementById(id);

  // ===== 担当者を自分にする =====
  // 位置(options[1])ではなく「<< 自分 >>」の選択肢を探して選ぶ
  const assignee = byId('issue_assigned_to_id');
  if (assignee) {
    const me = [...assignee.options].find((o) => o.text.trim().startsWith('<<'));
    if (me) assignee.value = me.value;
  }

  // 対象バージョンを設定する場合は、次の行のコメントを外す(選択肢の位置で指定)
  // byId('issue_fixed_version_id').options[1].selected = true;

  // ===== 子チケットの場合だけ、以降を実行 =====
  const parentId = new URLSearchParams(location.search).get('issue[parent_issue_id]');
  if (!parentId) return;

  // 親チケットの画面を取得する(URLは今開いているRedmineに合わせる)
  let parent;
  try {
    const res = await fetch(`${location.origin}/issues/${encodeURIComponent(parentId)}`, {
      credentials: 'same-origin',
    });
    if (!res.ok) {
      console.log(`${LOG} 親チケットの取得に失敗: ${res.status} ${res.statusText}`);
      return;
    }
    parent = new DOMParser().parseFromString(await res.text(), 'text/html');
  } catch (e) {
    console.log(`${LOG} 親チケットの取得に失敗:`, e);
    return;
  }

  // 親のタイトル(編集フォームの値 → 見出しの順に探す)
  const parentSubject =
    byId('issue_subject', parent)?.value?.trim() ||
    parent.querySelector('.issue .subject h3')?.textContent.trim() ||
    '';

  // 親のトラッカー名(編集フォームの選択値 → ページ見出し「ストーリー #123」の順に探す)
  const parentTracker =
    byId('issue_tracker_id', parent)?.selectedOptions[0]?.text.trim() ||
    (parent.querySelector('#content h2')?.textContent.trim().match(/^(.+?)\s+#\d+/) || [])[1] ||
    '';

  // ===== タイトルと説明を入力する =====
  // トラッカーを変えるとRedmineがフォームを再読み込みするが、
  // 入力済みの値は引き継がれるので、先に入れておく
  if (parentSubject) byId('issue_subject').value = `${parentSubject}:`;
  const description = byId('issue_description');
  if (description) description.value = DESCRIPTION_TEMPLATE;

  // ===== トラッカーを1階層下に変更する =====
  const childTracker = CHILD_TRACKER[parentTracker];
  const trackerSelect = byId('issue_tracker_id');
  const target =
    childTracker && trackerSelect && [...trackerSelect.options].find((o) => o.text.trim() === childTracker);

  console.log(`${LOG} 親: ${parentTracker || '(不明)'} → 子: ${target ? childTracker : '(変更なし)'}`);

  if (target && trackerSelect.value !== target.value) {
    trackerSelect.value = target.value;
    // Redmine標準の動き(トラッカーに合わせたフォームの再読み込み)を起こす
    trackerSelect.dispatchEvent(new Event('change', { bubbles: true }));
  }
})();
