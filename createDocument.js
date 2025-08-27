function createAndReplaceDocument() {
  const spreadsheetId = '1ZOOvmYeISSEqDD_UMikaXJzrVs0NnerPfW1g6TDiBLE';
  const sheetName = 'NotebookLM';
  const htmlColumn = 7; // G列
  const idColumn = 2; // B列
  const checkColumn = 4; // D列
  const linkColumn = 14; // N列
  const linkToInsertColumn = 3; // C列

  const folderId = '155cEenmQlRHStomkLskPNF3gA37vBKFh';
  
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  const folder = DriveApp.getFolderById(folderId);

  // ヘッダー行をスキップしてデータを処理
  for (let i = 1; i < values.length; i++) {
    // D列（checkColumn）がFALSEの行のみを対象とする
    if (values[i][checkColumn - 1] === false) {
      const originalHtmlContent = values[i][htmlColumn - 1];
      const id = values[i][idColumn - 1];
      const pageLink = values[i][linkToInsertColumn - 1];

      // HTMLコンテンツとIDが存在する場合に処理を実行
      if (originalHtmlContent && id) {
        const fileName = `${id}`;

        // 同じ名前の既存ファイルを検索し、見つかった場合は削除
        const files = folder.getFilesByName(fileName);
        while (files.hasNext()) {
          const file = files.next();
          file.setTrashed(true);
        }

        // 新しいGoogle Documentを作成
        const doc = DocumentApp.create(fileName);
        const body = doc.getBody();

        // C列のリンクをドキュメントの先頭に追加
        if (pageLink) {
          body.appendParagraph('オリジナルページ:');
          body.appendParagraph(pageLink).setLinkUrl(pageLink);
          body.appendParagraph(''); // 空行を追加
        }

        // HTMLコンテンツをテキストとして追加（HTMLタグを削除）
        const textContent = originalHtmlContent
          .replace(/<br\s*\/?>/gi, '\n')
          .replace(/<\/p>/gi, '\n\n')
          .replace(/<\/h[1-6]>/gi, '\n\n')
          .replace(/<li>/gi, '• ')
          .replace(/<\/li>/gi, '\n')
          .replace(/<[^>]*>/g, '')
          .replace(/&nbsp;/g, ' ')
          .replace(/&lt;/g, '<')
          .replace(/&gt;/g, '>')
          .replace(/&amp;/g, '&')
          .replace(/&quot;/g, '"')
          .replace(/&#39;/g, "'")
          .trim();

        // テキストを段落ごとに分割して追加
        const paragraphs = textContent.split(/\n\n+/);
        paragraphs.forEach(paragraph => {
          if (paragraph.trim()) {
            body.appendParagraph(paragraph.trim());
          }
        });

        // ドキュメントを指定のフォルダに移動
        const docFile = DriveApp.getFileById(doc.getId());
        folder.addFile(docFile);
        DriveApp.getRootFolder().removeFile(docFile);
        
        // 作成したドキュメントのリンクをN列（linkColumn）に貼り付け
        const fileUrl = doc.getUrl();
        sheet.getRange(i + 1, linkColumn).setValue(fileUrl);
      }
    }
  }
}