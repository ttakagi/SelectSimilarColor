// SelectSimilarColor.jsx
// 選択オブジェクトの色に近い色を持つオブジェクトをリアルタイムで選択

#target illustrator

(function () {

    // --- カラーユーティリティ ---

    // IllustratorのカラーオブジェクトをプレーンなJSオブジェクトにコピー
    function copyColor(color) {
        if (!color) return null;
        var t = color.typename;
        if (t === "CMYKColor")  return { type:"cmyk", c:color.cyan, m:color.magenta, y:color.yellow, k:color.black };
        if (t === "RGBColor")   return { type:"rgb",  r:color.red,  g:color.green,   b:color.blue };
        if (t === "GrayColor")  return { type:"gray", v:color.gray };
        if (t === "SpotColor")  return copyColor(color.spot.color);
        return null; // NoColor / GradientColor / PatternColor
    }

    // 表示用RGB（変換は表示のみに使用）
    function toDisplayRGB(co) {
        if (!co) return null;
        if (co.type === "rgb")  return { r: Math.round(co.r), g: Math.round(co.g), b: Math.round(co.b) };
        if (co.type === "gray") { var v = Math.round(255 * (1 - co.v / 100)); return { r:v, g:v, b:v }; }
        if (co.type === "cmyk") {
            var c = co.c/100, m = co.m/100, y = co.y/100, k = co.k/100;
            return {
                r: Math.round(255 * (1-c) * (1-k)),
                g: Math.round(255 * (1-m) * (1-k)),
                b: Math.round(255 * (1-y) * (1-k))
            };
        }
        return null;
    }

    // ネイティブ色空間で距離計算（CMYK同士はCMYK空間で比較）
    // 戻り値はRGB換算スケール（0〜441）で統一
    function colorDist(a, b) {
        if (!a || !b) return 9999;

        // 両方グレー→グレー空間
        if (a.type === "gray" && b.type === "gray") {
            return Math.abs(a.v - b.v) * 255 / 100;
        }

        // 両方CMYK→CMYK空間で比較（変換誤差なし）
        if (a.type === "cmyk" && b.type === "cmyk") {
            var dc = a.c - b.c, dm = a.m - b.m, dy = a.y - b.y, dk = a.k - b.k;
            // CMYK最大距離≈200、RGB最大距離≈441 → スケール合わせ
            return Math.sqrt(dc*dc + dm*dm + dy*dy + dk*dk) * 441 / 200;
        }

        // 両方RGB→RGB空間
        if (a.type === "rgb" && b.type === "rgb") {
            var dr = a.r - b.r, dg = a.g - b.g, db = a.b - b.b;
            return Math.sqrt(dr*dr + dg*dg + db*db);
        }

        // 色空間が異なる場合のみRGBに変換して比較
        var ra = toDisplayRGB(a), rb = toDisplayRGB(b);
        if (!ra || !rb) return 9999;
        var dr = ra.r-rb.r, dg = ra.g-rb.g, db = ra.b-rb.b;
        return Math.sqrt(dr*dr + dg*dg + db*db);
    }

    function getItemFillColor(item) {
        try {
            // CompoundPathItem は直接 fillColor を持たないため内部パスから取得
            if (item.typename === "CompoundPathItem") {
                if (item.pathItems && item.pathItems.length > 0) {
                    var c = item.pathItems[0].fillColor;
                    if (c && c.typename !== "NoColor") return copyColor(c);
                }
                return null;
            }
            // TextFrame はキャラクター属性から取得
            if (item.typename === "TextFrame") {
                try {
                    var c = item.textRange.characterAttributes.fillColor;
                    if (c && c.typename !== "NoColor") return copyColor(c);
                } catch (e) {}
                try {
                    var chars = item.textRange.characters;
                    if (chars.length > 0) {
                        var c = chars[0].characterAttributes.fillColor;
                        if (c && c.typename !== "NoColor") return copyColor(c);
                    }
                } catch (e) {}
                return null;
            }
            var c = item.fillColor;
            if (c && c.typename !== "NoColor") return copyColor(c);
        } catch (e) {}
        return null;
    }

    function getItemStrokeColor(item) {
        try {
            // CompoundPathItem も同様に内部パスから取得
            if (item.typename === "CompoundPathItem") {
                if (item.pathItems && item.pathItems.length > 0) {
                    var c = item.pathItems[0].strokeColor;
                    if (c && c.typename !== "NoColor") return copyColor(c);
                }
                return null;
            }
            // TextFrame はキャラクター属性から取得
            if (item.typename === "TextFrame") {
                try {
                    var c = item.textRange.characterAttributes.strokeColor;
                    if (c && c.typename !== "NoColor") return copyColor(c);
                } catch (e) {}
                return null;
            }
            var c = item.strokeColor;
            if (c && c.typename !== "NoColor") return copyColor(c);
        } catch (e) {}
        return null;
    }

    // --- アイテム収集（レイヤー再帰） ---

    function collectGroup(group, result) {
        try {
            for (var i = 0; i < group.pageItems.length; i++) {
                var item = group.pageItems[i];
                result.push(item);
                if (item.typename === "GroupItem") collectGroup(item, result);
            }
        } catch (e) {}
    }

    function collectLayer(layer, result) {
        try {
            for (var s = 0; s < layer.layers.length; s++) {
                collectLayer(layer.layers[s], result);
            }
            for (var i = 0; i < layer.pageItems.length; i++) {
                var item = layer.pageItems[i];
                result.push(item);
                if (item.typename === "GroupItem") collectGroup(item, result);
            }
        } catch (e) {}
    }

    function getAllItems(doc) {
        var r = [];
        for (var i = 0; i < doc.layers.length; i++) collectLayer(doc.layers[i], r);
        return r;
    }

    // --- 初期化 ---

    if (!app.documents.length) { alert("ドキュメントが開かれていません。"); return; }
    var doc = app.activeDocument;
    var sel = doc.selection;
    if (!sel || sel.length === 0) { alert("オブジェクトを選択してから実行してください。"); return; }

    var baseColor = null;
    var baseColorSource = "塗り";
    for (var si = 0; si < sel.length; si++) {
        baseColor = getItemFillColor(sel[si]);
        if (baseColor) { baseColorSource = "塗り"; break; }
    }
    if (!baseColor) {
        for (var si = 0; si < sel.length; si++) {
            baseColor = getItemStrokeColor(sel[si]);
            if (baseColor) { baseColorSource = "線"; break; }
        }
    }
    if (!baseColor) {
        alert("選択オブジェクトに有効な塗り・線の色が見つかりませんでした。\n（グラデーション・パターンは対象外です）");
        return;
    }

    var allItems = getAllItems(doc);

    // 表示用HEX
    function toHex(v) { var h = Math.max(0,Math.min(255,Math.round(v))).toString(16).toUpperCase(); return h.length<2?"0"+h:h; }
    var displayRGB = toDisplayRGB(baseColor);
    var hexStr = displayRGB
        ? "#" + toHex(displayRGB.r) + toHex(displayRGB.g) + toHex(displayRGB.b)
        : "(変換不可)";
    var colorSpaceLabel = baseColor.type === "cmyk" ? "CMYK" : baseColor.type === "rgb" ? "RGB" : "Gray";

    // item.visible はグループ内で誤ってfalseを返すことがある
    // → 親チェーンを辿りレイヤーの可視性で判定する
    function isLayerVisible(layer) {
        try {
            if (!layer.visible) return false;
            if (layer.parent && layer.parent.typename === "Layer") {
                return isLayerVisible(layer.parent);
            }
        } catch (e) {}
        return true;
    }

    function isItemVisible(item) {
        try {
            var p = item.parent;
            while (p) {
                if (p.typename === "Layer") return isLayerVisible(p);
                p = p.parent;
            }
        } catch (e) {}
        return true;
    }

    function isItemLocked(item) {
        try {
            var p = item;
            while (p && p.typename !== "Document") {
                if (p.locked) return true;
                p = p.parent;
            }
        } catch (e) {}
        return false;
    }

    // --- 選択実行 ---

    function applySelection(threshold, useFill, useStroke, inclLocked, inclHidden) {
        var matched = [];
        for (var i = 0; i < allItems.length; i++) {
            var item = allItems[i];
            if (!inclLocked && isItemLocked(item))   continue;
            if (!inclHidden && !isItemVisible(item)) continue;

            var hit = false;
            if (useFill && !hit) {
                var f = getItemFillColor(item);
                if (f && colorDist(baseColor, f) <= threshold) hit = true;
            }
            if (useStroke && !hit) {
                var s = getItemStrokeColor(item);
                if (s && colorDist(baseColor, s) <= threshold) hit = true;
            }
            if (hit) matched.push(item);
        }

        doc.selection = null;
        for (var j = 0; j < matched.length; j++) {
            try { matched[j].selected = true; } catch (e) {}
        }
        app.redraw();
        return matched.length;
    }

    // --- ダイアログ ---

    var dlg = new Window("dialog", "近似色を選択");
    dlg.orientation = "column";
    dlg.alignChildren = ["fill", "top"];
    dlg.spacing = 10;
    dlg.margins = 16;

    var previewGroup = dlg.add("group");
    previewGroup.add("statictext", undefined, "基準色:");
    var colorLabel = previewGroup.add("statictext", undefined, "");
    colorLabel.preferredSize.width = 280;
    if (displayRGB) {
        colorLabel.text = hexStr + "  ( R:" + displayRGB.r + "  G:" + displayRGB.g + "  B:" + displayRGB.b + " )  [" + colorSpaceLabel + " / " + baseColorSource + "]";
    }

    var targetPanel = dlg.add("panel", undefined, "比較対象");
    targetPanel.orientation = "row";
    targetPanel.margins = [10, 14, 10, 10];
    var chkFill   = targetPanel.add("checkbox", undefined, "塗り");
    var chkStroke = targetPanel.add("checkbox", undefined, "線");
    chkFill.value = true;

    var THRESH_MIN = 0, THRESH_MAX = 150, THRESH_DEFAULT = 30;

    var threshPanel = dlg.add("panel", undefined, "色の近さ（閾値）");
    threshPanel.orientation = "column";
    threshPanel.alignChildren = ["fill", "center"];
    threshPanel.margins = [10, 14, 10, 10];

    var sliderRow = threshPanel.add("group");
    sliderRow.orientation = "row";
    sliderRow.alignChildren = ["left", "center"];
    var sliderCtrl = sliderRow.add("slider", undefined, THRESH_DEFAULT, THRESH_MIN, THRESH_MAX);
    sliderCtrl.preferredSize.width = 200;
    var valField = sliderRow.add("edittext", undefined, String(THRESH_DEFAULT));
    valField.preferredSize.width = 45;
    valField.justify = "center";

    var scaleRow = threshPanel.add("group");
    scaleRow.orientation = "row";
    scaleRow.add("statictext", undefined, "完全一致 (0)");
    var sp = scaleRow.add("statictext", undefined, ""); sp.preferredSize.width = 110;
    scaleRow.add("statictext", undefined, "(150) 広め");

    var hintText  = threshPanel.add("statictext", undefined, ""); hintText.justify  = "center";
    var countText = threshPanel.add("statictext", undefined, ""); countText.justify = "center";

    // スキャン対象数を表示（デバッグ確認用）
    var scanInfo = threshPanel.add("statictext", undefined, "スキャン対象: " + allItems.length + " オブジェクト");
    scanInfo.justify = "center";

    function updateHint(val) {
        if (val === 0)       hintText.text = "完全一致のみ";
        else if (val <= 20)  hintText.text = "かなり近い色";
        else if (val <= 50)  hintText.text = "似た色";
        else if (val <= 100) hintText.text = "やや広め";
        else                 hintText.text = "広い範囲";
    }

    var optPanel = dlg.add("panel", undefined, "オプション");
    optPanel.orientation = "row";
    optPanel.margins = [10, 14, 10, 10];
    var chkLocked = optPanel.add("checkbox", undefined, "ロック済みも対象");
    var chkHidden = optPanel.add("checkbox", undefined, "非表示も対象");

    var btnGroup = dlg.add("group");
    btnGroup.orientation = "row";
    btnGroup.alignment = "right";
    var btnCancel = btnGroup.add("button", undefined, "キャンセル", { name: "cancel" });
    var btnOK     = btnGroup.add("button", undefined, "OK（選択を確定）", { name: "ok" });

    function doSelect() {
        var v = Math.round(sliderCtrl.value);
        updateHint(v);
        if (!chkFill.value && !chkStroke.value) {
            countText.text = "塗りか線を選択してください";
            return;
        }
        var n = applySelection(v, chkFill.value, chkStroke.value, chkLocked.value, chkHidden.value);
        countText.text = n > 0 ? n + " 個を選択中" : "一致なし";
    }

    sliderCtrl.onChanging = function () {
        var v = Math.round(sliderCtrl.value);
        valField.text = String(v);
        doSelect();
    };

    valField.onChange = function () {
        var v = parseInt(valField.text, 10);
        if (!isNaN(v)) {
            v = Math.max(THRESH_MIN, Math.min(THRESH_MAX, v));
            sliderCtrl.value = v;
            valField.text = String(v);
            doSelect();
        }
    };

    chkFill.onClick   = doSelect;
    chkStroke.onClick = doSelect;
    chkLocked.onClick = doSelect;
    chkHidden.onClick = doSelect;

    btnCancel.onClick = function () { doc.selection = null; dlg.close(); };
    btnOK.onClick     = function () { dlg.close(); };

    updateHint(THRESH_DEFAULT);
    doSelect();

    dlg.show();
})();
