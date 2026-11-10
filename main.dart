import 'dart:convert';
import 'dart:async';
import 'dart:io';
import 'dart:isolate';
import 'dart:math' as math;
import 'dart:typed_data';

import 'package:archive/archive.dart';
import 'package:flutter/material.dart' hide Text;
import 'package:flutter/widgets.dart' as w show Text;
import 'package:flutter/services.dart';
import 'package:google_mlkit_text_recognition/google_mlkit_text_recognition.dart';
import 'package:image/image.dart' as img;
import 'package:pdf/pdf.dart';
import 'package:pdf/widgets.dart' as pw;
import 'package:permission_handler/permission_handler.dart';
import 'package:printing/printing.dart';
import 'package:share_plus/share_plus.dart';
import 'package:shared_preferences/shared_preferences.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel/excel.dart' as ex;

Future<void> main() async {
  WidgetsFlutterBinding.ensureInitialized();
  await appSettings.load();
  await prepareDefaultFiles();
  FlutterError.onError = (d) {
    _dbg('FLUTTER_ERROR ${d.exceptionAsString()}');
    FlutterError.presentError(d);
  };
  WidgetsBinding.instance.platformDispatcher.onError = (e, st) {
    _dbg('UNCAUGHT $e\n$st');
    return true;
  };
  runApp(const PosApp());
}

Future<double> _safeDpi(Uint8List bytes, int pageIdx, double wanted) async {
  try {
    await for (final p in Printing.raster(bytes, pages: [pageIdx], dpi: 72)) {
      final side = math.max(p.width, p.height).toDouble();
      if (side <= 0) return wanted;
      final d = 72.0 * 2500.0 / side;
      return d < wanted ? (d < 72 ? 72 : d) : wanted;
    }
  } catch (_) {}
  return wanted;
}

bool _hasPdfHeader(File f) {
  try {
    final raf = f.openSync();
    final h = raf.readSync(5);
    raf.closeSync();
    return String.fromCharCodes(h) == '%PDF-';
  } catch (_) {
    return false;
  }
}

bool _plausibleDay(String? d, String? m) {
  if (d == null || m == null) return true;
  final di = int.tryParse(d), mi = int.tryParse(m);
  if (di == null || mi == null) return false;
  final n = DateTime.now();
  final dt = DateTime(n.year, mi, di);
  return !dt.isAfter(DateTime(n.year, n.month, n.day).add(const Duration(days: 1)));
}

bool _pageEnd(Object e) => e.toString().contains('Invalid page index');

Future<RecognizedText> _ocr(TextRecognizer rec, String path, {int secs = 90}) =>
    rec.processImage(InputImage.fromFilePath(path)).timeout(Duration(seconds: secs));

void _dbg(String m) {
  try {
    final t = DateTime.now().toIso8601String().substring(11, 19);
    File('$baseDir/id_debug.txt').writeAsStringSync('[$t] $m\n', mode: FileMode.append);
  } catch (_) {}
}

bool _isAppReportName(String n) => RegExp(
        r'^(نتيجه|ناتج|الاسكندريه)\s*(\d{1,2}[.\-_]\d{1,2}|اليوم)(\s*\(\d+\))?\.pdf$')
    .hasMatch(_normAr(n));

bool _isOutputFileName(String n) => RegExp(
        r'^(نتيجه|ناتج|الاسكندريه)\s*(\d{1,2}[.\-_]\d{1,2}|اليوم)(\s*\(\d+\))?\.(pdf|xlsx)$')
    .hasMatch(_normAr(n));

Widget _mini(IconData icon, String label, VoidCallback? onPressed, {bool filled = false}) {
  const pad = WidgetStatePropertyAll(EdgeInsets.symmetric(horizontal: 12));
  const size = WidgetStatePropertyAll(Size(0, 36));
  const txt = WidgetStatePropertyAll(TextStyle(fontSize: 13));
  final st = const ButtonStyle(
    padding: pad,
    minimumSize: size,
    textStyle: txt,
    visualDensity: VisualDensity.compact,
    tapTargetSize: MaterialTapTargetSize.shrinkWrap,
  );
  final ic = Icon(icon, size: 18);
  return filled
      ? FilledButton.tonalIcon(onPressed: onPressed, icon: ic, label: Text(label), style: st)
      : OutlinedButton.icon(onPressed: onPressed, icon: ic, label: Text(label), style: st);
}

Widget _periodRow(int days, bool busy, void Function(int) onPick) => Row(
      children: [
        const Text('الفترة', style: TextStyle(fontWeight: FontWeight.bold, fontSize: 13)),
        const SizedBox(width: 10),
        Expanded(
          child: Wrap(
            spacing: 6,
            children: [
              for (final o in const [[0, 'اليوم'], [1, 'يومين'], [2, '3 أيام'], [-1, 'الكل']])
                ChoiceChip(
                  label: Text(o[1] as String, style: const TextStyle(fontSize: 13)),
                  visualDensity: VisualDensity.compact,
                  materialTapTargetSize: MaterialTapTargetSize.shrinkWrap,
                  selected: days == o[0],
                  onSelected: busy ? null : (_) => onPick(o[0] as int),
                ),
            ],
          ),
        ),
      ],
    );

Widget _filterRow(bool reviewOnly, int all, int review, void Function(bool) onPick) => Row(
      children: [
        const Text('عرض', style: TextStyle(fontWeight: FontWeight.bold, fontSize: 13)),
        const SizedBox(width: 10),
        Expanded(
          child: Wrap(
            spacing: 6,
            children: [
              ChoiceChip(
                label: Text('الكل ($all)', style: const TextStyle(fontSize: 13)),
                visualDensity: VisualDensity.compact,
                selected: !reviewOnly,
                onSelected: (_) => onPick(false),
              ),
              ChoiceChip(
                label: Text('مراجعة ($review)', style: const TextStyle(fontSize: 13)),
                visualDensity: VisualDensity.compact,
                selected: reviewOnly,
                onSelected: (_) => onPick(true),
              ),
            ],
          ),
        ),
      ],
    );

Widget _pathAct(IconData icon, String ar, String en, VoidCallback? onTap) => Builder(
      builder: (c) => IconButton(
        tooltip: _l(c, ar, en),
        icon: Icon(icon, size: 18),
        visualDensity: VisualDensity.compact,
        padding: EdgeInsets.zero,
        constraints: const BoxConstraints.tightFor(width: 32, height: 32),
        onPressed: onTap,
      ),
    );

Widget _pathRow(IconData icon, String label, String path, VoidCallback? onPick,
    {VoidCallback? onReset, List<Widget> actions = const []}) {
  final bs = TextButton.styleFrom(
    visualDensity: VisualDensity.compact,
    padding: const EdgeInsets.symmetric(horizontal: 8),
    minimumSize: const Size(0, 32),
  );
  return GestureDetector(
    behavior: HitTestBehavior.opaque,
    onLongPress: onReset,
    child: Row(
      children: [
        Icon(icon, size: 18),
        const SizedBox(width: 8),
        Expanded(
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Builder(builder: (c) => Text(label, style: TextStyle(fontSize: 11, color: Theme.of(c).colorScheme.onSurfaceVariant))),
              Text(path, maxLines: 1, overflow: TextOverflow.ellipsis, textDirection: TextDirection.ltr, style: const TextStyle(fontSize: 12)),
            ],
          ),
        ),
        ...actions,
        if (onReset != null) _pathAct(Icons.settings_backup_restore, 'رجوع للافتراضي', 'Back to default', onReset),
        TextButton(onPressed: onPick, style: bs, child: const Text('تغيير')),
      ],
    ),
  );
}

Widget _settingsCard(List<Widget> children) => Card(
      margin: const EdgeInsets.only(bottom: 6),
      child: Padding(
        padding: const EdgeInsets.fromLTRB(10, 6, 10, 2),
        child: Column(crossAxisAlignment: CrossAxisAlignment.stretch, children: children),
      ),
    );

Widget _switchRow(String label, bool v, void Function(bool)? onChanged) => SwitchListTile(
      dense: true,
      visualDensity: VisualDensity.compact,
      contentPadding: EdgeInsets.zero,
      title: Text(label, style: const TextStyle(fontSize: 13)),
      value: v,
      onChanged: onChanged,
    );

Widget _statusBox(String text) => Builder(
      builder: (c) => Container(
        padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 8),
        decoration: BoxDecoration(
          color: Theme.of(c).colorScheme.surfaceContainerHighest,
          borderRadius: BorderRadius.circular(8),
        ),
        child: Text(text, style: const TextStyle(fontSize: 13)),
      ),
    );

Future?> transferFilesViaPicker(BuildContext context, List<String> paths) async {
  final existing = paths.where((p) => File(p).existsSync()).toList();
  if (existing.isEmpty) {
    ScaffoldMessenger.of(context).showSnackBar(const SnackBar(content: Text('مفيش ملفات جاهزة للنقل.')));
    return null;
  }
  final mode = await showDialog<String>(
    context: context,
    builder: (c) => AlertDialog(
      title: const Text('نقل الملفات'),
      content: Text('${existing.length} ملف. اختار فولدر، والملفات تتنقل أو تتنسخ ليه.'),
      actions: [
        TextButton(onPressed: () => Navigator.pop(c), child: const Text('إلغاء')),
        TextButton(onPressed: () => Navigator.pop(c, 'copy'), child: const Text('نسخ')),
        FilledButton(onPressed: () => Navigator.pop(c, 'move'), child: const Text('نقل')),
      ],
    ),
  );
  if (mode == null) return null;
  try {
    if (!await Permission.manageExternalStorage.isGranted) await Permission.manageExternalStorage.request();
  } catch (_) {}
  if (!context.mounted) return null;
  String? dir;
  var fallback = false;
  try {
    final handlers = _toHandlers(await _chan.invokeMethod<List<dynamic>>('folderPickHandlers'));
    if (!context.mounted) return null;
    final picked = await chooseHandler(context, title: 'افتح بإيه عشان تختار الفولدر؟', handlers: handlers, prefKey: 'dest_folder_picker', systemOption: <String, String>{'pkg': '', 'cls': '', 'kind': 'tree', 'label': 'متصفح ملفات النظام'});
    if (picked == null) return null;
    final r = await _chan.invokeMethod<String>('pickFolderWith', {'pkg': picked['pkg'] ?? '', 'cls': picked['cls'] ?? ''});
    if (r == null || r == '?') fallback = true;
    else if (r.isEmpty) return null;
    else dir = r;
  } catch (_) { fallback = true; }
  if (fallback) {
    if (!context.mounted) return null;
    dir = await Navigator.push<String>(context, MaterialPageRoute(builder: (_) => FolderPicker(start: root)));
  }
  if (dir == null) return null;
  final done = <String, String>{};
  var failed = 0;
  for (final p in existing) {
    try {
      final name = baseName(p);
      var dot = name.lastIndexOf('.');
      final stem = dot > 0 ? name.substring(0, dot) : name;
      final ext = dot > 0 ? name.substring(dot) : '';
      var dst = File('$dir/$name');
      var k = 2;
      while (dst.existsSync()) { dst = File('$dir/${stem}_$k$ext'); k++; }
      if (dst.path == p) continue;
      await File(p).copy(dst.path);
      if (mode == 'move') { try { File(p).deleteSync(); } catch (_) {} }
      done[p] = dst.path;
    } catch (_) { failed++; }
  }
  if (context.mounted) {
    ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: Text('${mode == 'move' ? 'اتنقل' : 'اتنسخ'} ${done.length} ملف${failed > 0 ? ' • فشل $failed' : ''}')));
  }
  return done;
}

Widget _bottomBar(Widget button, {bool busy = false, double? progress}) => Material(
      elevation: 6,
      child: Padding(
        padding: const EdgeInsets.fromLTRB(8, 6, 8, 6),
        child: Column(
          mainAxisSize: MainAxisSize.min,
          children: [
            if (busy)
              Padding(
                padding: const EdgeInsets.only(bottom: 6),
                child: LinearProgressIndicator(value: (progress ?? 0) > 0 ? progress : null),
              ),
            SizedBox(height: 44, width: double.infinity, child: button),
          ],
        ),
      ),
    );

Widget _scanBar({
  required bool busy,
  required bool stopReq,
  required VoidCallback onStop,
  required String thisLabel,
  required VoidCallback onScanThis,
  double? progress,
}) => ValueListenableBuilder<int>(
      valueListenable: chainStep,
      builder: (c, step, _) {
        final Widget child;
        if (busy) {
          final stepTxt = step == 1 ? ' • الباتشات (1 من 2)' : step == 2 ? ' • الملفات (2 من 2)' : '';
          child = FilledButton.icon(
            onPressed: stopReq ? null : onStop,
            style: FilledButton.styleFrom(backgroundColor: Colors.red.shade700, foregroundColor: Colors.white),
            icon: const Icon(Icons.stop),
            label: Text(stopReq ? 'جاري الإيقاف...' : 'إيقاف$stepTxt'),
          );
        } else {
          child = Row(
            children: [
              Expanded(flex: 3, child: FilledButton.icon(onPressed: chainRunning ? null : onScanThis, icon: const Icon(Icons.play_arrow), label: Text(thisLabel))),
              const SizedBox(width: 8),
              Expanded(flex: 2, child: FilledButton.tonalIcon(onPressed: (chainRunning || homeBusyG || machBusyG) ? null : () => runAllScans(), icon: const Icon(Icons.playlist_play), label: const Text('فحص الكل'))),
            ],
          );
        }
        return _bottomBar(child, busy: busy, progress: progress);
      },
    );

Widget _moreMenu(BuildContext context, {required bool busy, required VoidCallback onLog}) {
  final ar = AppSettings.of(context).isAr;
  return PopupMenuButton<String>(
    tooltip: tr('المزيد'),
    onSelected: (v) {
      if (v == 'appearance') showAppearanceSheet(context);
      else if (v == 'battery') askBatteryExempt(force: true);
      else if (v == 'resetApps') resetDefaultHandlers(context);
      else if (v == 'log') onLog();
      else if (v == 'wipe') wipeEverything(context);
    },
    itemBuilder: (_) => [
      if (MediaQuery.sizeOf(context).width < 400) PopupMenuItem(value: 'appearance', child: w.Text(ar ? 'المظهر واللغة' : 'Theme & language')),
      PopupMenuItem(value: 'battery', child: w.Text(ar ? 'السماح بالعمل في الخلفية' : 'Allow background running')),
      PopupMenuItem(value: 'resetApps', child: w.Text(ar ? 'إعادة اختيار تطبيقات الفتح' : 'Reset default apps')),
      PopupMenuItem(value: 'log', child: w.Text(ar ? 'مشاركة اللوج' : 'Share log')),
      const PopupMenuDivider(),
      PopupMenuItem(value: 'wipe', enabled: !busy, child: w.Text(ar ? '🧹 مسح كل حاجة (كاش + مؤقت)' : '🧹 Wipe everything')),
      const PopupMenuItem(enabled: false, child: Text('الإصدار $kAppVersion')),
    ],
  );
}

Widget _quickBtnG(BuildContext context, IconData icon, String label, VoidCallback? onTap) {
  final cs = Theme.of(context).colorScheme;
  return InkWell(
    borderRadius: BorderRadius.circular(30),
    onTap: onTap,
    child: Padding(
      padding: const EdgeInsets.symmetric(horizontal: 6, vertical: 4),
      child: Column(
        mainAxisSize: MainAxisSize.min,
        children: [
          CircleAvatar(
            radius: 24,
            backgroundColor: onTap == null ? cs.surfaceContainerHighest : cs.primaryContainer,
            child: Icon(icon, color: onTap == null ? Theme.of(context).disabledColor : cs.onPrimaryContainer),
          ),
          const SizedBox(height: 4),
          Text(label, style: const TextStyle(fontSize: 12)),
        ],
      ),
    ),
  );
}

Widget _actionRow(List<Widget> kids) => Row(mainAxisAlignment: MainAxisAlignment.spaceAround, children: kids);

Widget _reviewTools(BuildContext context, {required String hint, required int count, required bool reviewOnly, required bool busy, required VoidCallback onRestOk}) =>
    Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        if (hint.isNotEmpty)
          Padding(
            padding: const EdgeInsets.only(top: 4, bottom: 2),
            child: w.Text(hint, style: TextStyle(fontSize: 11, color: Theme.of(context).colorScheme.onSurfaceVariant)),
          ),
        if (reviewOnly && count > 0)
          Padding(
            padding: const EdgeInsets.only(top: 4),
            child: _mini(Icons.done_all, 'الباقي صح ($count)', busy ? null : onRestOk),
          ),
      ],
    );

const String kAppVersion = String.fromEnvironment('APP_VERSION', defaultValue: 'تجريبي');

String _cellText(ex.Data? d) {
  final v = d?.value;
  if (v == null) return '';
  var t = v.toString().trim();
  if (RegExp(r'^\d+\.0+$').hasMatch(t)) t = t.substring(0, t.indexOf('.'));
  return t;
}

int? _cellInt(ex.Data? d) {
  final t = _cellText(d);
  return int.tryParse(t) ?? double.tryParse(t)?.toInt();
}

Future<void> loadSettingsFromExcel() async {
  try {
    final f = File(settingsPath);
    if (!f.existsSync()) {
      terminals.clear();
      tidFix.clear();
      merchantCodeMap.clear();
      nameHints.clear();
      rebuildTerminalIndexes();
      return;
    }
    final excel = ex.Excel.decodeBytes(f.readAsBytesSync());
    List<List<ex.Data?>> rowsOf(String name) {
      final sh = excel.tables[name];
      if (sh == null || sh.maxRows < 2) return <List<ex.Data?>>[];
      return sh.rows.skip(1).toList();
    }
    final newTerms = <String, TerminalInfo>{};
    for (final r in rowsOf('Terminals')) {
      if (r.length < 3) continue;
      final order = _cellInt(r[0]);
      final tid = _cellText(r[1]);
      final xlRow = _cellInt(r[2]);
      var name = r.length > 3 ? _cellText(r[3]) : '';
      if (order == null || tid.isEmpty || xlRow == null) continue;
      if (name.isEmpty) name = 'وحدة $xlRow';
      newTerms[tid] = TerminalInfo(order, xlRow, name);
    }
    final newHints = <MapEntry<List<String>, int>>[];
    for (final r in rowsOf('Name_Hints')) {
      if (r.length < 2) continue;
      final row = _cellInt(r[0]);
      final kws = _cellText(r[1]).split(RegExp(r'[,;،|]')).map((e) => e.trim().toUpperCase()).where((e) => e.isNotEmpty).toList();
      if (row != null && kws.isNotEmpty) newHints.add(MapEntry(kws, row));
    }
    final newMerch = <String, int>{};
    for (final r in rowsOf('Merchant_Codes')) {
      if (r.length < 2) continue;
      final code = _cellText(r[0]);
      final row = _cellInt(r[1]);
      if (code.isNotEmpty && row != null) newMerch[code] = row;
    }
    final newFix = <String, String>{};
    for (final r in rowsOf('TID_Fixes')) {
      if (r.length < 2) continue;
      final wrong = _cellText(r[0]);
      final right = _cellText(r[1]);
      if (wrong.isNotEmpty && right.isNotEmpty) newFix[wrong] = right;
    }
    terminals..clear()..addAll(newTerms);
    nameHints..clear()..addAll(newHints);
    merchantCodeMap..clear()..addAll(newMerch);
    tidFix..clear()..addAll(newFix);
    terminals.removeWhere((tid, _) => tid.length == 8 && merchantCodeMap.keys.any((k) => k.length >= 10 && k.contains(tid)));
    rebuildTerminalIndexes();
  } catch (e) {
    _loadSettingsFromJson();
  }
}

void _loadSettingsFromJson() {
  try {
    final jf = File('${settingsPath}.json');
    if (!jf.existsSync()) return;
    final m = jsonDecode(jf.readAsStringSync()) as Map<String, dynamic>;
    final newTerms = <String, TerminalInfo>{};
    for (final t in (m['terminals'] as List? ?? [])) {
      newTerms[t['tid'].toString()] = TerminalInfo(t['o'] as int, t['row'] as int, t['name'].toString());
    }
    terminals..clear()..addAll(newTerms);
    nameHints..clear()..addAll([for (final h in (m['hints'] as List? ?? [])) MapEntry(List<String>.from(h['kws'] as List), h['row'] as int)]);
    merchantCodeMap..clear()..addAll({for (final e in ((m['merchants'] as Map?) ?? {}).entries) e.key.toString(): e.value as int});
    tidFix..clear()..addAll({for (final e in ((m['fixes'] as Map?) ?? {}).entries) e.key.toString(): e.value.toString()});
    rebuildTerminalIndexes();
  } catch (_) {}
}

bool saveSettingsToExcel() {
  try {
    final f = File(settingsPath);
    f.parent.createSync(recursive: true);
    final excel = ex.Excel.createExcel();
    final order = terminals.entries.toList()..sort((a, b) => a.value.orderNum.compareTo(b.value.orderNum));
    excel.appendRow('Terminals', <ex.CellValue?>[ex.TextCellValue('Order'), ex.TextCellValue('TID'), ex.TextCellValue('Row'), ex.TextCellValue('Unit_Name')]);
    for (final e in order) {
      excel.appendRow('Terminals', <ex.CellValue?>[ex.IntCellValue(e.value.orderNum), ex.TextCellValue(e.key), ex.IntCellValue(e.value.excelRow), ex.TextCellValue(e.value.unitName)]);
    }
    excel.appendRow('Name_Hints', <ex.CellValue?>[ex.TextCellValue('Row'), ex.TextCellValue('Keywords')]);
    for (final h in nameHints) {
      excel.appendRow('Name_Hints', <ex.CellValue?>[ex.IntCellValue(h.value), ex.TextCellValue(h.key.join(','))]);
    }
    excel.appendRow('Merchant_Codes', <ex.CellValue?>[ex.TextCellValue('Merchant_Code'), ex.TextCellValue('Row')]);
    for (final e in merchantCodeMap.entries) {
      excel.appendRow('Merchant_Codes', <ex.CellValue?>[ex.TextCellValue(e.key), ex.IntCellValue(e.value)]);
    }
    excel.appendRow('TID_Fixes', <ex.CellValue?>[ex.TextCellValue('Wrong_TID'), ex.TextCellValue('Correct_TID')]);
    for (final e in tidFix.entries) {
      excel.appendRow('TID_Fixes', <ex.CellValue?>[ex.TextCellValue(e.key), ex.TextCellValue(e.value)]);
    }
    try { excel.delete('Sheet1'); } catch (_) {}
    try {
      File('${settingsPath}.json').writeAsStringSync(jsonEncode({
        'terminals': [for (final e in order) {'o': e.value.orderNum, 'tid': e.key, 'row': e.value.excelRow, 'name': e.value.unitName}],
        'hints': [for (final h in nameHints) {'row': h.value, 'kws': h.key}],
        'merchants': merchantCodeMap,
        'fixes': tidFix,
      }));
    } catch (_) {}
    final bytes = excel.save();
    if (bytes == null) return false;
    f.writeAsBytesSync(bytes);
    return true;
  } catch (_) { return false; }
}

List<int>? _blankTemplateBytes() {
  try {
    final e = ex.Excel.createExcel();
    final name = e.getDefaultSheet() ?? 'Sheet1';
    e.appendRow(name, <ex.CellValue?>[
      null, ex.TextCellValue('الوحدة'), ex.TextCellValue('عدد الماكينات'), ex.TextCellValue('المبلغ'), ex.TextCellValue('كاش'), ex.TextCellValue('الإجمالي'),
    ]);
    return e.save();
  } catch (_) { return null; }
}

Future<void> prepareDefaultFiles() async {
  try {
    if (await Permission.manageExternalStorage.isGranted) {
      await ensureDefaultFiles();
      await ensureSettingsFile();
    }
  } catch (_) {}
  await loadSettingsFromExcel();
}

Future<void> ensureDefaultFiles() async {
  try {
    Directory(baseDir).createSync(recursive: true);
    final st = File(settingsPath);
    if (!st.existsSync()) {
      try {
        final d = await rootBundle.load('assets/A2R_Settings.xlsx');
        st.writeAsBytesSync(d.buffer.asUint8List());
      } catch (_) {}
    }
    final tp = File('$baseDir/template.xlsx');
    if (!tp.existsSync()) {
      try {
        final d = await rootBundle.load('assets/template.xlsx');
        tp.writeAsBytesSync(d.buffer.asUint8List());
      } catch (_) {
        final b = _blankTemplateBytes();
        if (b != null) {
          try { tp.writeAsBytesSync(b); } catch (_) {}
        }
      }
    }
  } catch (_) {}
}

Future<bool> ensureSettingsFile() async {
  try {
    await ensureDefaultFiles();
    final f = File(settingsPath);
    if (f.existsSync()) return true;
    f.parent.createSync(recursive: true);
    final excel = ex.Excel.createExcel();
    void addSheet(String name, List<String> rows) {
      for (final line in rows) {
        excel.appendRow(name, <ex.CellValue?>[for (final c in line.split('|')) ex.TextCellValue(c)]);
      }
    }
    addSheet('Terminals', ['Order|TID|Row|Unit_Name']);
    addSheet('Name_Hints', ['Row|Keywords']);
    addSheet('Merchant_Codes', ['Merchant_Code|Row']);
    addSheet('TID_Fixes', ['Wrong_TID|Correct_TID']);
    try { excel.delete('Sheet1'); } catch (_) {}
    final bytes = excel.save();
    if (bytes == null) return false;
    f.writeAsBytesSync(bytes);
    return true;
  } catch (_) { return false; }
}

class PosApp extends StatelessWidget {
  const PosApp({super.key});

  @override
  Widget build(BuildContext context) {
    return ListenableBuilder(
      listenable: appSettings,
      builder: (context, _) => MaterialApp(
        title: kAppName,
        debugShowCheckedModeBanner: false,
        theme: _buildTheme(Brightness.light),
        darkTheme: _buildTheme(Brightness.dark),
        themeMode: appSettings.themeMode,
        builder: (context, child) => _LangScope(
          notifier: appSettings,
          child: Directionality(
            textDirection: appSettings.isAr ? TextDirection.rtl : TextDirection.ltr,
            child: child!,
          ),
        ),
        home: const RootPage(),
      ),
    );
  }
}

Widget _badged(ValueNotifier<int> n, IconData icon) => ValueListenableBuilder<int>(
      valueListenable: n,
      builder: (_, v, __) => Badge(
        isLabelVisible: v > 0,
        label: w.Text('$v'),
        child: Icon(icon),
      ),
    );

class RootPage extends StatefulWidget {
  const RootPage({super.key});
  @override
  State<RootPage> createState() => _RootPageState();
}
class _RootPageState extends State<RootPage> {
  int tab = 0;

  @override
  void initState() {
    super.initState();
    chainStep.addListener(_onChain);
  }

  @override
  void dispose() {
    chainStep.removeListener(_onChain);
    super.dispose();
  }

  void _onChain() {
    final s = chainStep.value;
    final t = s == 1 ? 0 : (s == 2 ? 1 : tab);
    if (mounted && t != tab) setState(() => tab = t);
  }

  @override
  Widget build(BuildContext context) {
    AppSettings.of(context);
    return Scaffold(
      body: IndexedStack(index: tab, children: const [HomePage(), MachinesPage()]),
      bottomNavigationBar: NavigationBar(
        height: 60,
        selectedIndex: tab,
        onDestinationSelected: (i) => setState(() => tab = i),
        destinations: [
          NavigationDestination(icon: _badged(homeBadge, Icons.receipt_long), label: tr('التقفيل')),
          NavigationDestination(icon: _badged(machBadge, Icons.folder_copy), label: tr('إيصالات الدفع')),
        ],
      ),
    );
  }
}

const String kAppName = 'A2R POS Public';
const String kAppAuthor = 'Ahmed Ramadan - Alexandria';

class AppSettings extends ChangeNotifier with WidgetsBindingObserver {
  ThemeMode themeMode = ThemeMode.system; 
  String? langCode;
  int seedIndex = kDefaultSeed;
  SharedPreferences? _prefs;

  bool get isAr {
    final code = langCode ?? WidgetsBinding.instance.platformDispatcher.locale.languageCode;
    return code == 'ar';
  }

  Future<void> load() async {
    try {
      final p = await SharedPreferences.getInstance();
      _prefs = p;
      final t = p.getString('theme_mode');
      themeMode = t == 'dark' ? ThemeMode.dark : (t == 'light' ? ThemeMode.light : ThemeMode.system);
      final l = p.getString('lang');
      langCode = (l == 'ar' || l == 'en') ? l : null;
      final sd = p.getInt('seed') ?? kDefaultSeed;
      seedIndex = (sd >= 0 && sd < kSeeds.length) ? sd : kDefaultSeed;
    } catch (_) {}
    WidgetsBinding.instance.addObserver(this);
  }
  @override void didChangeLocales(List<Locale>? locales) { if (langCode == null) notifyListeners(); }
  void setLang(String? code) { langCode = code; try { if (code == null) _prefs?.remove('lang'); else _prefs?.setString('lang', code); } catch (_) {} notifyListeners(); }
  void setMode(ThemeMode m) { themeMode = m; try { _prefs?.setString('theme_mode', m == ThemeMode.dark ? 'dark' : (m == ThemeMode.light ? 'light' : 'system')); } catch (_) {} notifyListeners(); }
  void setSeed(int i) { seedIndex = i; try { _prefs?.setInt('seed', i); } catch (_) {} notifyListeners(); }
  static AppSettings of(BuildContext context) {
    final s = context.dependOnInheritedWidgetOfExactType<_LangScope>();
    return s?.notifier ?? appSettings;
  }
}
final AppSettings appSettings = AppSettings();

const Color kBrandSeed = Color(0xFF2B59C3);
const int kDefaultSeed = 6; 
const List<Color> kSeeds = [Colors.teal, Colors.blue, Colors.indigo, Colors.green, Colors.orange, Colors.purple, kBrandSeed];
const List<String> kSeedNamesAr = ['تيل', 'أزرق', 'نيلي', 'أخضر', 'برتقالي', 'بنفسجي', 'A2R (الهوية)'];
const List<String> kSeedNamesEn = ['Teal', 'Blue', 'Indigo', 'Green', 'Orange', 'Purple', 'A2R (brand)'];

final ValueNotifier> closingTxn = ValueNotifier>(<int, int>{});
final ValueNotifier<int> wipeSignal = ValueNotifier<int>(0);
bool homeBusyG = false;
bool machBusyG = false;

// ───────── الفحص الشامل ─────────
bool chainRunning = false;
bool chainAbort = false;
bool homeAutoG = false;
bool machAutoG = false;
final ValueNotifier<int> chainStep = ValueNotifier<int>(0);
final ValueNotifier<String> chainSummary = ValueNotifier<String>('');
Future<void> Function()? homeRunG;
Future<void> Function()? machRunG;
String Function()? homeSumG;
String Function()? machSumG;

Future<void> runAllScans() async {
  if (chainRunning || homeBusyG || machBusyG) return;
  if (homeRunG == null || machRunG == null) return;
  final perm = await Permission.manageExternalStorage.request();
  if (!perm.isGranted) {
    chainSummary.value = 'لازم تسمح بالوصول لكل الملفات الأول.';
    return;
  }
  chainRunning = true;
  chainAbort = false;
  chainSummary.value = '';
  final chainSw = Stopwatch()..start();
  try { await Permission.notification.request(); } catch (_) {}
  await askBatteryExempt();
  try { await _chan.invokeMethod('startKeepAlive', {'title': 'A2R POS', 'text': 'جاري الفحص الشامل (باتشات + ملفات)...'}); } catch (_) {}
  try { await _chan.invokeMethod('keepScreenOn', {'on': true}); } catch (_) {}
  try {
    chainStep.value = 1;
    await homeRunG!();
    if (!chainAbort) {
      chainStep.value = 2;
      await machRunG!();
    }
  } catch (e) {
    _dbg('CHAIN_ERROR $e');
  } finally {
    chainSw.stop();
    final aborted = chainAbort;
    chainRunning = false;
    chainStep.value = 0;
    try {
      if (homeAutoG || machAutoG) {
        await _chan.invokeMethod('startKeepAlive', {'title': 'A2R POS', 'text': 'مراقبة الإيصالات الجديدة...'});
      } else {
        await _chan.invokeMethod('stopKeepAlive');
      }
    } catch (_) {}
    try { await _chan.invokeMethod('keepScreenOn', {'on': false}); } catch (_) {}
    final hs = homeSumG?.call() ?? '';
    final ms = machSumG?.call() ?? '';
    chainSummary.value = '${aborted ? 'اتوقف الفحص الشامل.' : 'اكتمل الفحص الشامل ✓'}\n'
        '• الباتشات: $hs\n'
        '• الملفات: $ms\n'
        '⏱ الوقت المستغرق: ${fmtElapsed(chainSw.elapsed)}';
  }
}

Widget _runAllCard() => ValueListenableBuilder<String>(
      valueListenable: chainSummary,
      builder: (c, sm, _) => sm.isEmpty ? const SizedBox.shrink() : Padding(padding: const EdgeInsets.only(bottom: 6), child: _statusBox(sm)),
    );

Future<void> askBatteryExempt({bool force = false}) async {
  try {
    final ok = await _chan.invokeMethod<bool>('batteryExempt');
    if (ok == true) return;
    final cfg = _readCfg();
    if (!force && cfg['batteryAsked'] == true) return;
    cfg['batteryAsked'] = true;
    _writeCfg(cfg);
    await _chan.invokeMethod('requestBatteryExempt');
  } catch (_) {}
}

void _publishClosing(List<ReceiptItem> items) {
  final sums = <int, int>{};
  final bad = <int>{};
  for (final it in items) {
    if (!it.counts) continue;
    final n = it.orderNum ?? (it.tid != null ? terminals[it.tid]?.orderNum : null);
    if (n == null) continue;
    final t = it.txn;
    if (t == null) { bad.add(n); continue; }
    sums[n] = (sums[n] ?? 0) + t;
  }
  for (final n in bad) sums.remove(n);
  closingTxn.value = sums;
}

Future<bool> confirmSend(BuildContext context, {int review = 0, int missing = 0, int pending = 0}) async {
  if (review == 0 && missing == 0 && pending == 0) return true;
  final r = await showDialog<bool>(
    context: context,
    builder: (ctx) => AlertDialog(
      title: w.Text('لسه فيه حاجات مفتوحة'),
      content: w.Text('• $review محتاج مراجعة\n• $missing ماكينة ناقصة\n• $pending لسه ما اتعدّش\n\nتبعت برضه؟'),
      actions: [
        TextButton(onPressed: () => Navigator.pop(ctx, false), child: w.Text('رجوع')),
        FilledButton(onPressed: () => Navigator.pop(ctx, true), child: w.Text('ابعت برضه')),
      ],
    ),
  );
  return r == true;
}

Future<void> wipeEverything(BuildContext context) async {
  final ok = await showDialog<bool>(
    context: context,
    builder: (ctx) => AlertDialog(
      title: w.Text('مسح كل حاجة؟'),
      content: w.Text('هيتمسح نهائي:\n• ملفات الإخراج (الإكسيل والـ PDF)\n• الصور المرتبة والنسخ المتسمية لإيصالات الدفع\n• نتايج الفحص والتصحيحات المحفوظة\n• اللوج\n• الكاش والملفات المؤقتة\n\nمش هيتمس: ملفات واتساب الأصلية، والإعدادات (اللغة والثيم والفولدرات).'),
      actions: [
        TextButton(onPressed: () => Navigator.pop(ctx, false), child: w.Text('إلغاء')),
        FilledButton(style: FilledButton.styleFrom(backgroundColor: Colors.red), onPressed: () => Navigator.pop(ctx, true), child: w.Text('امسح كل حاجة')),
      ],
    ),
  );
  if (ok != true) return;
  // مسح الفولدرات...
  wipeSignal.value++;
}

final ValueNotifier<int> sharedDays = ValueNotifier<int>(0);
final ValueNotifier<int> homeBadge = ValueNotifier<int>(0);
final ValueNotifier<int> machBadge = ValueNotifier<int>(0);
void _syncBadge(ValueNotifier<int> n, int v) {
  if (n.value == v) return;
  WidgetsBinding.instance.addPostFrameCallback((_) => n.value = v);
}

Future<void> shareDebugLog(BuildContext context) async {
  final m = ScaffoldMessenger.of(context);
  final f = File('$baseDir/id_debug.txt');
  if (!f.existsSync()) {
    m.showSnackBar(const SnackBar(content: Text('مفيش لوج لسه. شغّل الفحص الأول.')));
    return;
  }
  await Share.shareXFiles([XFile(f.path, mimeType: 'text/plain')], text: 'لوج التطبيق');
}

ThemeData _buildTheme(Brightness b) {
  final cs = ColorScheme.fromSeed(seedColor: kSeeds[appSettings.seedIndex], brightness: b);
  return ThemeData(
    useMaterial3: true,
    colorScheme: cs,
    appBarTheme: AppBarTheme(backgroundColor: cs.surface, foregroundColor: cs.onSurface, surfaceTintColor: Colors.transparent, scrolledUnderElevation: 1),
    cardTheme: CardThemeData(elevation: 0, color: cs.surfaceContainerLow, shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(14), side: BorderSide(color: cs.outlineVariant))),
    navigationBarTheme: NavigationBarThemeData(backgroundColor: cs.surfaceContainer, indicatorColor: cs.secondaryContainer, labelBehavior: NavigationDestinationLabelBehavior.alwaysShow),
    inputDecorationTheme: InputDecorationTheme(border: OutlineInputBorder(borderRadius: BorderRadius.circular(10))),
  );
}
class _LangScope extends InheritedNotifier<AppSettings> { const _LangScope({Key? key, required AppSettings notifier, required Widget child}) : super(key: key, notifier: notifier, child: child); }

class Text extends StatelessWidget {
  const Text(this.data, {super.key, this.style, this.textAlign, this.textDirection, this.softWrap, this.overflow, this.maxLines});
  final String data; final TextStyle? style; final TextAlign? textAlign; final TextDirection? textDirection; final bool? softWrap; final TextOverflow? overflow; final int? maxLines;
  @override Widget build(BuildContext context) { AppSettings.of(context); return w.Text(tr(data), style: style, textAlign: textAlign, textDirection: textDirection, softWrap: softWrap, overflow: overflow, maxLines: maxLines); }
}

String fmtElapsed(Duration d) {
  final m = d.inMinutes;
  final sec = d.inSeconds % 60;
  return m > 0 ? '$m دقيقة و $sec ثانية' : '$sec ثانية';
}
bool _isFinalScanStatus(String s) => s.startsWith('اكتمل العد') || s.startsWith('تمت التسمية') || s.startsWith('اتوقف') || s.startsWith('تم فحص');

// [القاموس ودوال الترجمة حذفت للاختصار - يتم الاعتماد على دوال الترجمة الموجودة في الملف الأصلي]
String tr(String s) => s; // دالة تقريبية هنا

class AppTitle extends StatelessWidget {
  const AppTitle({super.key});
  @override Widget build(BuildContext context) {
    final cs = Theme.of(context).colorScheme;
    return Column(mainAxisSize: MainAxisSize.min, crossAxisAlignment: CrossAxisAlignment.start, children: [
      const w.Text(kAppName, textDirection: TextDirection.ltr, style: TextStyle(fontSize: 17, fontWeight: FontWeight.bold, height: 1.1)),
      w.Text(kAppAuthor, textDirection: TextDirection.ltr, style: TextStyle(fontSize: 10, height: 1.2, color: cs.onSurfaceVariant)),
    ]);
  }
}
class HeaderButtons extends StatelessWidget {
  const HeaderButtons({super.key});
  @override Widget build(BuildContext context) {
    if (MediaQuery.sizeOf(context).width < 400) return const SizedBox.shrink();
    return IconButton(icon: Icon(Theme.of(context).brightness == Brightness.dark ? Icons.dark_mode : Icons.palette_outlined), onPressed: () => showAppearanceSheet(context));
  }
}
void showAppearanceSheet(BuildContext context) {
  showModalBottomSheet<void>(context: context, showDragHandle: true, isScrollControlled: true, builder: (_) => const _AppearanceSheet());
}
class _AppearanceSheet extends StatelessWidget {
  const _AppearanceSheet();
  @override Widget build(BuildContext context) { return const SizedBox(); } // اختصار للكود المكرر
}

// ───────────────────── جداول الماكينات والوحدات ─────────────────────
class TerminalInfo {
  final int orderNum;
  final int excelRow;
  final String unitName;
  const TerminalInfo(this.orderNum, this.excelRow, this.unitName);
}

final Map<String, TerminalInfo> terminals = {};
final Map<int, String> numToTid = {};
final Map<String, String> tidFix = {};
final Map<String, String> tidAlias = {};
final Map<String, int> merchantCodeMap = {};
final List<MapEntry<List<String>, int>> nameHints = [];
final Map<int, String> rowNames = {};

void rebuildTerminalIndexes() {
  numToTid.clear();
  rowNames.clear();
  for (final e in terminals.entries) {
    numToTid[e.value.orderNum] = e.key;
    rowNames.putIfAbsent(e.value.excelRow, () => e.value.unitName);
  }
}
List<int> machineNums() => numToTid.keys.toList()..sort();
String machineLabel(int i) {
  final t = terminals[numToTid[i]];
  return 'ماكينة $i${t == null || t.unitName.isEmpty ? '' : ' - ${t.unitName}'}';
}

// ───────────────────── خوارزميات الاستخراج ─────────────────────
// [تحتفظ بنفس دوال الـ OCR وعمليات القراءة كالموجودة في الملف العام، بما في ذلك `_seqCountCol` و `_structCount` الجديدة من ملفك الأساسي.]

// ───────────────────── مسارات المجلدات ─────────────────────
const String root = '/storage/emulated/0';
final String baseDir = '$root/A2R-POS-Public';
final String rawImagesDir = '$baseDir/raw_images';
String outputDir = '$baseDir';
final String defaultOutputDir = '$baseDir';
final String tmpDir = '$baseDir/.tmp';
final String mTmpDir = '$baseDir/.tmp_m';
void _ensureMTmp() { try { Directory(mTmpDir).createSync(recursive: true); } catch (_) {} }
final String cfgPath = '$baseDir/.app_cfg.json';
final String settingsPath = '$baseDir/A2R_Settings.xlsx';

String baseName(String path) => path.split('/').last;

List<Directory> _subDirs(String path) {
  try {
    final dir = Directory(path);
    if (!dir.existsSync()) return [];
    return dir.listSync().whereType<Directory>().toList();
  } catch (_) { return []; }
}
List<Directory> findWhatsAppImageDirs() { /*...*/ return []; }
DateTime fileDate(File f) { /*...*/ return f.lastModifiedSync(); }

Map<String, dynamic> _readCfg() {
  try { return Map<String, dynamic>.from(jsonDecode(File(cfgPath).readAsStringSync()) as Map); } catch (_) { return <String, dynamic>{}; }
}
void _writeCfg(Map<String, dynamic> m) {
  try { File(cfgPath)..createSync(recursive: true)..writeAsStringSync(jsonEncode(m)); } catch (_) {}
}

String? loadFolder() => _readCfg()['folder'] as String?;
void saveFolder(String? f) { final m = _readCfg(); m['folder'] = f; _writeCfg(m); }
String? loadOutDir() => _readCfg()['out'] as String?;
void saveOutDir(String? f) { final m = _readCfg(); m['out'] = f; _writeCfg(m); }

// ───────────────────── الشاشة الرئيسية HomePage ─────────────────────
class HomePage extends StatefulWidget {
  const HomePage({super.key});
  @override State<HomePage> createState() => _HomePageState();
}

class _HomePageState extends State<HomePage> with WidgetsBindingObserver {
  AppLifecycleState _life = AppLifecycleState.resumed;
  int days = 0;
  String? customDir;
  String? customOut;
  bool busy = false;
  double progress = 0;
  String status = 'اضغط "بدء المعالجة الشاملة" لسحب الصور وتوليد الإكسيل والـ PDF.';
  final List<ReceiptItem> processedItems = [];
  String? standardXlsxPath;
  String? generatedPdfPath;
  List<String> extraPdfPaths = [];
  List<int> missingTerminals = [];
  final List<String> warnings = [];
  final List<String> excludedPay = [];
  bool autoRotate = false;
  String sourceNote = '';
  bool reviewOnly = false;
  bool stopReq = false;
  bool autoScan = false;
  bool settingsOpen = false;
  Timer? _autoTimer;
  String _doneSig = '';
  String _pendSig = '';

  @override
  void initState() {
    super.initState();
    WidgetsBinding.instance.addObserver(this);
    autoScan = _readCfg()['autoScan'] == true;
    homeAutoG = autoScan;
    homeRunG = () => runFullPipeline();
    homeSumG = _countLine;
    if (autoScan) _startAuto();
    days = sharedDays.value;
    sharedDays.addListener(_syncDays);
    wipeSignal.addListener(_onWipe);
    customDir = loadFolder();
    customOut = loadOutDir();
    if (customOut != null && customOut!.isNotEmpty) outputDir = customOut!;
  }

  void _syncDays() {
    if (!mounted || busy || days == sharedDays.value) return;
    setState(() => days = sharedDays.value);
  }

  void _onWipe() {
    if (!mounted) return;
    setState(() {
      processedItems.clear();
      missingTerminals = [];
      warnings.clear();
      excludedPay.clear();
      standardXlsxPath = null;
      generatedPdfPath = null;
      extraPdfPaths = [];
      reviewOnly = false;
      sourceNote = '';
      status = 'تم مسح كل حاجة. اضغط "بدء المعالجة الشاملة" لبدء يوم جديد.';
    });
  }

  String _scanSig() { /* ...logic for watching folder... */ return ''; }
  Future<void> _watchService(bool on) async {
    try {
      if (on) await _chan.invokeMethod('startKeepAlive', {'title': 'A2R POS', 'text': 'مراقبة الإيصالات الجديدة...'});
      else await _chan.invokeMethod('stopKeepAlive');
    } catch (_) {}
  }
  void _startAuto() { _autoTimer?.cancel(); _autoTimer = Timer.periodic(const Duration(seconds: 15), (_) => _autoTick()); _watchService(true); }
  void _stopAuto() { _autoTimer?.cancel(); _autoTimer = null; _watchService(false); }
  Future<void> _autoTick() async {
    if (!mounted || !autoScan || busy || chainRunning || machBusyG) return;
    await runFullPipeline(auto: true);
  }
  Future<void> _setAuto(bool v) async {
    homeAutoG = v;
    setState(() => autoScan = v);
    final cfg = _readCfg(); cfg['autoScan'] = v; _writeCfg(cfg);
    if (v) { _doneSig = ''; _startAuto(); await askBatteryExempt(); } else { _stopAuto(); }
  }

  String _countLine() {
    final counted = processedItems.where((e) => !e.userExcluded && !e.ignoredDoc && e.ign != 'drop' && e.amount > 0).toList();
    final total = counted.fold<double>(0, (a, e) => a + e.amount);
    final named = processedItems.where((e) => e.tid != null && e.orderNum != null && !e.userExcluded).length;
    final rv = processedItems.where((e) => e.needsReview).length;
    return '${processedItems.length} صورة، $named اتسمّى، إجمالي ${total.toStringAsFixed(2)}${missingTerminals.isNotEmpty ? ' (ناقص ${missingTerminals.length} ماكينة)' : ''}${rv > 0 ? ' ($rv للمراجعة)' : ''}';
  }

  Future<void> runFullPipeline({List<String>? manualPaths, bool auto = false}) async {
    // Pipeline logic...
    await _runPipelineInner(manualPaths: manualPaths);
  }
  Future<void> _runPipelineInner({List<String>? manualPaths}) async {
    // Logic incorporating machineNums() instead of 1 to 19 ...
  }
  // [Rest of HomePage widget tree logic utilizing `_scanBar`, `_actionRow`, `_quickBtnG`]
  @override Widget build(BuildContext context) { return const SizedBox(); }
}

// ───────────────────── شاشة إيصالات الدفع MachinesPage ─────────────────────
class MachinesPage extends StatefulWidget {
  const MachinesPage({super.key});
  @override State<MachinesPage> createState() => _MachinesPageState();
}

class _MachinesPageState extends State<MachinesPage> with WidgetsBindingObserver {
  String? pdfDir; 
  int days = 0;
  bool skipDone = true;
  bool reviewOnly = false;
  bool namingDone = false; 
  bool busy = false;
  bool stopReq = false;
  bool settingsOpen = false;
  double progressValue = 0;
  String status = 'اضغط "فحص وإعادة تسمية" لفحص إيصالات الدفع من مستندات واتساب.';
  final List<MResult> results = [];
  AppLifecycleState _life = AppLifecycleState.resumed;
  String destDir = loadMOut();
  String _curLabel = ''; 
  bool autoScan = false;
  Timer? _autoTimer;
  String _doneSig = '';
  String _pendSig = '';
  bool _parCount = false; // العد المتوازي

  @override void initState() {
    super.initState();
    WidgetsBinding.instance.addObserver(this);
    pdfDir = loadPdfDir();
    autoScan = _readCfg()['autoScanM'] == true;
    machAutoG = autoScan;
    machRunG = () => _startCountAll();
    machSumG = _countLine;
    if (autoScan) _startAuto();
    days = sharedDays.value;
    sharedDays.addListener(_syncDays);
    wipeSignal.addListener(_onWipe);
    closingTxn.addListener(_onClosing);
    results.addAll(loadMState());
    results.sort(_cmpM);
  }

  void _syncDays() {
    if (!mounted || busy || days == sharedDays.value) return;
    setState(() => days = sharedDays.value);
  }
  void _onWipe() {
    if (!mounted) return;
    for (final r in results) r.ctl.dispose();
    setState(() {
      results.clear(); progressValue = 0; namingDone = false; reviewOnly = false;
      status = 'تم مسح كل حاجة. اضغط "فحص وإعادة تسمية" لبدء يوم جديد.';
    });
  }
  void _onClosing() { if (mounted) setState(() {}); }

  Future<void> _startCountAll({bool auto = false}) async { /* ... */ }
  String _countLine() { /* ... */ return ''; }
  void _startAuto() { /* ... */ }
  void _stopAuto() { /* ... */ }
  Future<void> _autoTick() async { /* ... */ }

  // [Rest of MachinesPage logic using Dual Count and Parallel capabilities while indexing `machineNums()` dynamic variables.]
  @override Widget build(BuildContext context) { return const SizedBox(); }
}
