/**
 * ** 概要 **
 * 見た目を維持しながら塗り付きのオープンパスを自動処理するスクリプトです
 *
 * ** 注意 **
 * 1. アピアランスについて
 * 本スクリプトではパスを閉じる際にアンカーのハンドルを操作します
 * ハンドルに左右されるアピアランス（「パスの変形」項目など）がついている場合，見た目が変わってしまいます
 * 不安な場合は事前にアピアランス分割することをおすすめします
 *
 * 2. シンボルアイテムやプラグインアイテムについて
 * シンボルや複合シェイプやブレンドといったアイテム内にある塗り付きのオープンパスは本スクリプトでは検出できません
 * それらをチェックしたい場合は，各アイテムを拡張するなどしてから本スクリプトを実行してください
 *
 * 3. 複合パス内にある線付き塗りオープンパスについて
 * 複合パス内では，動作12（線と塗りを別のパスに分割）ができません
 * 複合パス内の線付き塗りオープンパスは，スクリプト実行後に個別で対応してください
 * 該当パスがあった場合は，スクリプト処理中に該当パスを選択してアラートします
 *
 * 4. その他
 * - レイヤーとオブジェクトのロックは自動で解除します
 * - ガイドには手をつけないので，オープンパスの数を確認する際は注意してください
 * - 本スクリプトは「見た目を変えずに」塗り付きのオープンパスを処理するものです
 *   線を足しながらオープンパスを閉じたい場合は「ナイフツールで囲む」といった方法が有用です
 * - 本スクリプトは自己責任においてご利用ください
 * - 本スクリプトに不具合や改善すべき点がございましたら，作成元まで連絡いただけると幸いです
 *
 * ** 動作 **
 * -- レイヤーに関する処理 --
 * 1. 全レイヤーのロックを解除します
 * 2. 空レイヤーを削除します
 * 3. 非表示レイヤーがあれば，ダイアログを出して削除するか確認します
 * 4. ダイアログで許可があれば非表示レイヤーを削除します
 *
 * -- オブジェクトに関する前処理 --
 * 5. 全オブジェクトのロックを解除します
 * 6. 非表示アイテムがあれば，ダイアログを出して削除するか確認します
 * 7. ダイアログで許可があれば非表示オブジェクトを削除します
 * 8. 孤立点があれば，ダイアログを出して削除するか確認します
 * 9. ダイアログで許可があれば孤立点を削除します
 *
 * -- 塗り付きのオープンパスの処理 --
 * 10. 塗りオープンパスのうち，直線パス（塗りの面積はないが塗り色がついている）を処理します（**詳細**の1）
 * 11. 塗りオープンパスのうち，線がついていないものはそのままパスを閉じます（**詳細**の2）
 * 12. 線がついている塗りオープンパスは，線と塗りを別々のパスにして，塗りのパスのみを閉じます（**詳細**の3）
 *
 * -- 結果のアラート --
 * 13. 扱った塗りオープンパスの個数をアラートして終了します
 *
 * ** 詳細 **
 * 1. 動作10（塗りがついている直線の処理）
 * 塗りはあるが線はないヘアラインパスと，塗りも線もついているパスで処理が分かれます
 * 前者の場合はパスを削除します。後者の場合は塗りを消去します
 * 処理の前にはダイアログを出して確認します。不許可の場合は何もしません
 * ※直線パスでもクローズパスの場合は本スクリプトの処理の対象外になります
 *
 * 2. 動作11（線がついていない塗りオープンパスの処理）
 * 塗り部分の見た目を変えないように始点アンカーと終点アンカーをつなげて，パスを閉じます
 * 処理としては「パスの連結」(Ctrl(⌘)+J)に近いです（なお「パスの連結」では見た目が変わる場合があります）
 *
 * 3. 動作12（線がついている塗りオープンパスの処理）
 * 前述の方法では，パスに線がついている場合はアンカー同士をつなげた分の線が足されてしまいます
 * そうならないように，線と塗りを別のパスに分割して，塗りのパスにだけ前述の処理をしてパスを閉じます
 * 線と塗りの分割では，パスを複製し，片方は塗りを消去，もう片方は線を消去しています
 */

// @target 'illustrator'

/**
 * 「再表示しない」ボックスがついた確認ダイアログのクラス
 */
var Confirmer = (function () {
  var Confirmer = function () {
    // newをつけ忘れた場合に備えて
    if (!(this instanceof Confirmer)) {
      return new Confirmer();
    }
    this.shouldConfirm = true; // 確認ダイアログを表示するかどうか
    this.result = false; // 確認の結果
  };

  /**
   * 確認の結果を返す。必要に応じて確認ダイアログを表示する
   * @param {string} body ダイアログの本文
   * @param {string} [title="重要な確認事項があります"] （省略可）ダイアログのタイトル
   * @param {string} [yesBtnTxt="はい"] （省略可）承諾ボタンのテキスト
   * @param {string} [noBtnTxt="いいえ"] （省略可）拒否ボタンのテキスト
   * @returns {boolean} 確認の結果
   */
  Confirmer.prototype.confirm = function (body, title, yesBtnTxt, noBtnTxt) {
    if (this.shouldConfirm) {
      // デフォルト引数
      if (title === undefined) title = '重要な確認事項があります';
      if (yesBtnTxt === undefined) yesBtnTxt = 'はい';
      if (noBtnTxt === undefined) noBtnTxt = 'いいえ';

      var dlg = new Window('dialog', title);
      dlg.add('statictext', undefined, body, { multiline: true });

      var btnGrp = dlg.add('group');
      var yesBtn = btnGrp.add('button', undefined, yesBtnTxt);
      yesBtn.onClick = function () {
        dlg.close(1); // dlg.show()の返り値が1になる
      };
      var noBtn = btnGrp.add('button', undefined, noBtnTxt);
      noBtn.onClick = function () {
        dlg.close(0); // dlg.show()の返り値が0になる
      };

      var checkbox = dlg.add(
        'checkbox',
        undefined,
        'このダイアログを再表示しない'
      );
      checkbox.value = false; // チェックボックスのデフォルトの値

      var result = dlg.show();
      if (result === 2) return false; // ダイアログが閉じられた場合の処理
      this.result = result;
      this.shouldConfirm = !checkbox.value;
    }
    return this.result;
  };

  return Confirmer;
})();

/**
 * メインプロセス
 */
(function () {
  if (app.documents.length <= 0) {
    alert('ファイルを開いてください。');
    return;
  }
  var doc = app.activeDocument;
  var result = ''; // 最後にアラートする文章

  // レイヤーに関する処理
  var hiddenLays = handleLayers(doc); // 返り値は非表示レイヤーのリスト

  // すべてのオブジェクトのロックを解除
  app.executeMenuCommand('unlockAll');
  unselectAll();

  // 確認ダイアログで許可があれば，非表示オブジェクトを削除
  var hiddenItems = removeHiddenItems(); // 返り値は非表示オブジェクトのリスト

  // 塗り付きのオープンパスに関する処理
  var paths = doc.pathItems;
  var lenPath = paths.length;
  var fopCount = 0;
  var cfmStrayPoint = new Confirmer();
  var cfmFilledLine = new Confirmer();
  for (var i = lenPath - 1; i >= 0; i--) {
    var currentPath = paths[i];
    // 孤立点について
    if (currentPath.pathPoints.length <= 1) {
      removeStrayPoint(currentPath, cfmStrayPoint);
      continue;
    }
    // ガイド以外の塗りがついているオープンパスについて
    if (!currentPath.guides && currentPath.filled && !currentPath.closed) {
      fopCount++;
      if (isLinear(currentPath)) {
        fixFilledLine(currentPath, cfmFilledLine);
      } else {
        fixFilledOpenPath(currentPath);
      }
    }
  }
  if (fopCount > 0) {
    result += fopCount + '個の塗り付きのオープンパスを処理しました。';
  }
  unselectAll();

  // 削除しなかった非表示オブジェクトを，非表示に戻す
  for (var i = 0, len = hiddenItems.length; i < len; i++) {
    hiddenItems[i].hidden = true;
  }

  // 削除しなかった非表示レイヤーを，非表示に戻す
  for (var i = 0, len = hiddenLays.length; i < len; i++) {
    hiddenLays[i].visible = false;
  }

  // 結果を表示
  if (result === '') result = '塗り付きのオープンパスはありませんでした。';
  alert(result);
})();

/**
 * すべてのレイヤーを表示＆ロック解除 + 空レイヤーを削除 +
 * 許可があれば非表示レイヤーを削除 + 削除しなかった非表示レイヤーを返す
 * @param {Document} docObj ドキュメント
 * @returns {Array<Layer>} 削除しなかった非表示レイヤーのリスト
 */
function handleLayers(docObj) {
  var result = []; // 非表示レイヤーを入れるリスト
  var confirmer = new Confirmer();
  (function rec(docObj) {
    var lays = docObj.layers;
    var len = lays.length;
    for (var i = len - 1; i >= 0; i--) {
      var currentLay = lays[i];
      // レイヤーのロックを解除
      currentLay.locked = false;
      // 空レイヤーを削除
      if (isEmptyLayer(currentLay)) {
        currentLay.visible = true;
        currentLay.remove();
        continue;
      }
      // 非表示レイヤーに関する処理
      if (!currentLay.visible) {
        currentLay.visible = true; // レイヤーを表示
        var shouldRemove = confirmer.confirm(
          'レイヤー「' +
            currentLay.name +
            '」は非表示レイヤーです。オブジェクトごと削除しますか？',
          '非表示のレイヤーがあります',
          '削除する',
          '削除しない'
        );
        if (shouldRemove) {
          currentLay.remove();
          continue;
        } else {
          // 削除しないものはリストに格納
          result.push(currentLay);
        }
      }
      rec(currentLay); // サブレイヤーのために再帰呼び出し
    }
  })(docObj);
  return result;
}

/**
 * レイヤーが完全に空かどうかを返す
 * @param {Layer} layObj レイヤー
 * @returns {boolean}
 */
function isEmptyLayer(layObj) {
  return layObj.pageItems.length === 0 && layObj.layers.length === 0;
}

/**
 * すべての選択を解除する
 */
function unselectAll() {
  app.selection = null;
}

/**
 * すべてのオブジェクトを表示 + 非表示オブジェクトを確認の上で削除 +
 * 削除しなかった非表示オブジェクトを返す
 * @returns {Array<PageItem>} 削除しなかった非表示オブジェクトのリスト
 */
function removeHiddenItems() {
  var result = []; // 非表示オブジェクトを入れるリスト
  var confirmer = new Confirmer();

  app.executeMenuCommand('showAll'); // すべてのオブジェクトを表示
  var hiddenItems = app.selection;
  var len = hiddenItems.length;
  for (var i = len - 1; i >= 0; i--) {
    var currentItem = hiddenItems[i];
    if (confirmer.shouldConfirm) {
      app.selection = currentItem;
      app.redraw(); // 選択を描画
    }
    var shouldRemove = confirmer.confirm(
      '現在選択されているオブジェクトは非表示のオブジェクトです。削除しますか？',
      '非表示のオブジェクトがあります',
      '削除する',
      '削除しない'
    );
    if (shouldRemove) {
      currentItem.remove();
    } else {
      result.push(currentItem);
    }
  }
  unselectAll();
  return result;
}

/**
 * 孤立点を削除する
 * @param {PathItem} pathObj 孤立点
 * @param {Confirmer} confirmer
 */
function removeStrayPoint(pathObj, confirmer) {
  if (confirmer.shouldConfirm) {
    app.selection = pathObj;
    app.redraw(); // 選択を描画
  }
  var shouldRemove = confirmer.confirm(
    '現在選択されているパスは孤立点です。孤立点を削除しますか？',
    '孤立点があります',
    '削除する',
    '削除しない'
  );
  if (shouldRemove) {
    pathObj.remove();
  }
}

/**
 * 一直線のパスかどうかを返す
 * @param {PathItem} pathObj パス
 * @returns {boolean}
 */
function isLinear(pathObj) {
  // パスの面積がほぼ0かどうかで判定
  return Math.round(Math.abs(pathObj.area) * 100000) === 0;
}

/**
 * 塗り付きの直線パスを処理する
 * @param {PathItem} pathObj 塗り付きの直線パス
 * @param {Confirmer} confirmer
 */
function fixFilledLine(pathObj, confirmer) {
  if (confirmer.shouldConfirm) {
    app.selection = pathObj;
    app.redraw(); // 選択を描画
  }
  if (pathObj.stroked) {
    var shouldErase = confirmer.confirm(
      '現在選択されている直線パスには，塗りがついています。塗りを消去しますか？',
      '直線パスに塗りがついています',
      '消去する',
      '消去しない'
    );
    if (shouldErase) {
      pathObj.filled = false;
    }
  } else {
    var shouldRemove = confirmer.confirm(
      '現在選択されている直線は，線なし塗り付きのヘアラインパスです。パスを削除しますか？',
      'ヘアラインパスがあります',
      '削除する',
      '削除しない'
    );
    if (shouldRemove) {
      pathObj.remove();
    }
  }
}

/**
 * 塗り付きのオープンパスを処理する
 * @param {PathItem} pathObj 塗りオープンパス
 */
function fixFilledOpenPath(pathObj) {
  if (!pathObj.stroked) {
    // 線がない場合は塗りを維持しながら閉じる
    closePath(pathObj);
  } else {
    if (isCompounded(pathObj)) {
      app.selection = pathObj;
      app.redraw(); // 選択を描画
      alert(
        '現在選択されているパスは複合パス内の塗り付きオープンパスです。\n' +
          'このスクリプトでは自動対処しないので，後で個別対応をお願いします。'
      );
    } else {
      // パスを複製し，片方は線を消去して閉じる，もう片方は塗りを消去して線のみにする
      var copied = pathObj.duplicate();
      pathObj.stroked = false;
      closePath(pathObj);
      copied.filled = false;
    }
  }
}

/**
 * 塗りの形を維持しながらパスを閉じる
 * @param {PathItem} pathObj オープンパス
 */
function closePath(pathObj) {
  var anchors = pathObj.pathPoints;
  var firstAnchor = anchors[0];
  var lastAnchor = anchors[anchors.length - 1];

  pathObj.closed = true; // ハンドルはそのままでパスを閉じる
  if (
    // 始点アンカーと終点アンカーが重なっているとき
    firstAnchor.anchor[0] === lastAnchor.anchor[0] &&
    firstAnchor.anchor[1] === lastAnchor.anchor[1]
  ) {
    // 終点アンカーを削除しても見た目が変わらないように始点アンカーの内向きハンドルを操作
    firstAnchor.leftDirection = lastAnchor.leftDirection;
    lastAnchor.remove();
  } else {
    // 始点アンカーの内向きハンドルを削除
    firstAnchor.leftDirection = firstAnchor.anchor;
    // 終点アンカーの外向きハンドルを削除
    lastAnchor.rightDirection = lastAnchor.anchor;
  }
}

/**
 * 複合パスの一部かどうかを返す
 * @param {PathItem} pathObj パス
 * @returns {boolean}
 */
function isCompounded(pathObj) {
  var parentObj = pathObj.parent;
  if (parentObj.typename === 'Layer') {
    return false;
  } else if (parentObj.typename === 'CompoundPathItem') {
    return true;
  } else {
    // パスのparentがレイヤーでも複合パスでもない，つまりグループの場合は再帰呼び出し
    return isCompounded(parentObj);
  }
}
