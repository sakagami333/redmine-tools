/* ==========================================================
   チケットの表:題名にマウスを乗せると、全文を表示する
   (design.css で題名を「…」で省略している場合の補助)

   View customize の設定:
     パスのパターン       :(空欄)
     プロジェクトのパターン:(空欄)
     挿入位置             :全ページのヘッダ
     種別                 :JavaScript
   ========================================================== */
document.addEventListener('DOMContentLoaded', function () {
  document.querySelectorAll('table.list.issues td.subject').forEach(function (td) {
    td.title = td.textContent.replace(/\s+/g, ' ').trim();
  });
});
