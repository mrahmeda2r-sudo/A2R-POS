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

// بيختار DPI بحيث الضلع الأطول ما يزيدش عن ~2500 بكسل
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
        r'^(نتيجه|ناتج)\s*(\d{1,2}[.\-_]\d{1,2}|اليوم)(\s*\(\d+\))?\.pdf$')
    .hasMatch(_normAr(n));

bool _isOutputFileName(String n) => RegExp(
        r'^(نتيجه|ناتج)\s*(\d{1,2}[.\-_]\d{1,2}|اليوم)(\s*\(\d+\))?\.(pdf|xlsx)$')
    .hasMatch(_normAr(n));

String fmtElapsed(Duration d) {
  final m = d.inMinutes;
  final sec = d.inSeconds % 60;
  return m > 0 ? '$m دقيقة و $sec ثانية' : '$sec ثانية';
}

bool _isFinalScanStatus(String s) =>
    s.startsWith('اكتمل العد') || s.startsWith('تمت التسمية') || s.startsWith('اتوقف') || s.startsWith('تم فحص');

// ───────── واجهة المستخدم الموحدة والمدمجة ─────────
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
              for (final o in const [
                [0, 'اليوم'],
                [1, 'يومين'],
                [2, '3 أيام'],
                [-1, 'الكل'],
              ])
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
                materialTapTargetSize: MaterialTapTargetSize.shrinkWrap,
                selected: !reviewOnly,
                onSelected: (_) => onPick(false),
              ),
              ChoiceChip(
                label: Text('مراجعة ($review)', style: const TextStyle(fontSize: 13)),
                visualDensity: VisualDensity.compact,
                materialTapTargetSize: MaterialTapTargetSize.shrinkWrap,
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
    tapTargetSize: MaterialTapTargetSize.shrinkWrap,
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
              Builder(
                builder: (c) => Text(label,
                    style: TextStyle(fontSize: 11, color: Theme.of(c).colorScheme.onSurfaceVariant)),
              ),
              Text(path,
                  maxLines: 1,
                  overflow: TextOverflow.ellipsis,
                  textDirection: TextDirection.ltr,
                  style: const TextStyle(fontSize: 12)),
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
    ScaffoldMessenger.of(context).showSnackBar(
      const SnackBar(content: Text('مفيش ملفات جاهزة للنقل.')),
    );
    return null;
  }
  final mode = await showDialog<String>(
    context: context,
    builder: (c) => AlertDialog(
      title: const Text('نقل الملفات'),
      content: Text('${existing.length} ملف. اختار فولدر من التليفون، وبعدين الملفات تتنقل أو تتنسخ ليه.'),
      actions: [
        TextButton(onPressed: () => Navigator.pop(c), child: const Text('إلغاء')),
        TextButton(onPressed: () => Navigator.pop(c, 'copy'), child: const Text('نسخ')),
        FilledButton(onPressed: () => Navigator.pop(c, 'move'), child: const Text('نقل')),
      ],
    ),
  );
  if (mode == null) return null;
  try {
    if (!await Permission.manageExternalStorage.isGranted) {
      await Permission.manageExternalStorage.request();
    }
  } catch (_) {}
  if (!context.mounted) return null;
  String? dir = await Navigator.push<String>(
    context,
    MaterialPageRoute(builder: (_) => FolderPicker(start: root)),
  );
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
      while (dst.existsSync()) {
        dst = File('$dir/${stem}_$k$ext');
        k++;
      }
      if (dst.path == p) continue;
      await File(p).copy(dst.path);
      if (mode == 'move') {
        try {
          File(p).deleteSync();
        } catch (_) {}
      }
      done[p] = dst.path;
    } catch (_) {
      failed++;
    }
  }
  if (context.mounted) {
    ScaffoldMessenger.of(context).showSnackBar(SnackBar(
      content: Text(
          '${mode == 'move' ? 'اتنقل' : 'اتنسخ'} ${done.length} ملف${failed > 0 ? ' • فشل $failed' : ''}'),
    ));
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
}) =>
    ValueListenableBuilder<int>(
      valueListenable: chainStep,
      builder: (c, step, _) {
        final Widget child;
        if (busy) {
          final stepTxt = step == 1
              ? ' • الباتشات (1 من 2)'
              : step == 2
                  ? ' • الملفات (2 من 2)'
                  : '';
          child = FilledButton.icon(
            onPressed: stopReq ? null : onStop,
            style: FilledButton.styleFrom(
                backgroundColor: Colors.red.shade700, foregroundColor: Colors.white),
            icon: const Icon(Icons.stop),
            label: Text(stopReq ? 'جاري الإيقاف...' : 'إيقاف$stepTxt'),
          );
        } else {
          child = Row(
            children: [
              Expanded(
                flex: 3,
                child: FilledButton.icon(
                  onPressed: chainRunning ? null : onScanThis,
                  icon: const Icon(Icons.play_arrow),
                  label: Text(thisLabel),
                ),
              ),
              const SizedBox(width: 8),
              Expanded(
                flex: 2,
                child: FilledButton.tonalIcon(
                  onPressed: (chainRunning || homeBusyG || machBusyG) ? null : () => runAllScans(),
                  icon: const Icon(Icons.playlist_play),
                  label: const Text('فحص الكل'),
                ),
              ),
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
      if (v == 'appearance') {
        showAppearanceSheet(context);
      } else if (v == 'battery') {
        askBatteryExempt(force: true);
      } else if (v == 'resetApps') {
        resetDefaultHandlers(context);
      } else if (v == 'log') {
        onLog();
      } else if (v == 'settings') {
        ensureSettingsFile().then((ok) {
          if (!context.mounted) return;
          if (ok) {
            openFileStatic(context, settingsPath, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
          }
        });
      } else if (v == 'wipe') {
        wipeEverything(context);
      }
    },
    itemBuilder: (_) => [
      if (MediaQuery.sizeOf(context).width < 400)
        PopupMenuItem(value: 'appearance', child: w.Text(ar ? 'المظهر واللغة' : 'Theme & language')),
      PopupMenuItem(
          value: 'battery', child: w.Text(ar ? 'السماح بالعمل في الخلفية' : 'Allow background running')),
      PopupMenuItem(
          value: 'resetApps', child: w.Text(ar ? 'إعادة اختيار تطبيقات الفتح' : 'Reset default apps')),
      PopupMenuItem(value: 'log', child: w.Text(ar ? 'مشاركة اللوج' : 'Share log')),
      PopupMenuItem(value: 'settings', child: const Text('إعدادات الماكينات (إكسيل)')),
      const PopupMenuDivider(),
      PopupMenuItem(
          value: 'wipe',
          enabled: !busy,
          child: w.Text(ar ? '🧹 مسح كل حاجة (كاش + مؤقت)' : '🧹 Wipe everything (cache + temp)')),
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

Widget _actionRow(List<Widget> kids) =>
    Row(mainAxisAlignment: MainAxisAlignment.spaceAround, children: kids);

Widget _reviewTools(BuildContext context,
        {required String hint,
        required int count,
        required bool reviewOnly,
        required bool busy,
        required VoidCallback onRestOk}) =>
    Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        if (hint.isNotEmpty)
          Padding(
            padding: const EdgeInsets.only(top: 4, bottom: 2),
            child: w.Text(hint,
                style: TextStyle(fontSize: 11, color: Theme.of(context).colorScheme.onSurfaceVariant)),
          ),
        if (reviewOnly && count > 0)
          Padding(
            padding: const EdgeInsets.only(top: 4),
            child: _mini(Icons.done_all, 'الباقي صح ($count)', busy ? null : onRestOk),
          ),
      ],
    );

const String kAppVersion = String.fromEnvironment('APP_VERSION', defaultValue: 'عام');

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
      final kws = _cellText(r[1])
          .split(RegExp(r'[,;،|]'))
          .map((e) => e.trim().toUpperCase())
          .where((e) => e.isNotEmpty)
          .toList();
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

    terminals
      ..clear()
      ..addAll(newTerms);
    nameHints
      ..clear()
      ..addAll(newHints);
    merchantCodeMap
      ..clear()
      ..addAll(newMerch);
    tidFix
      ..clear()
      ..addAll(newFix);
    terminals.removeWhere((tid, _) =>
        tid.length == 8 && merchantCodeMap.keys.any((k) => k.length >= 10 && k.contains(tid)));
    rebuildTerminalIndexes();
  } catch (e) {
    _dbg('Excel load error: $e');
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
      newTerms[t['tid'].toString()] =
          TerminalInfo(t['o'] as int, t['row'] as int, t['name'].toString());
    }
    terminals
      ..clear()
      ..addAll(newTerms);
    nameHints
      ..clear()
      ..addAll([
        for (final h in (m['hints'] as List? ?? []))
          MapEntry(List<String>.from(h['kws'] as List), h['row'] as int)
      ]);
    merchantCodeMap
      ..clear()
      ..addAll({
        for (final e in ((m['merchants'] as Map?) ?? {}).entries) e.key.toString(): e.value as int
      });
    tidFix
      ..clear()
      ..addAll({
        for (final e in ((m['fixes'] as Map?) ?? {}).entries) e.key.toString(): e.value.toString()
      });
    rebuildTerminalIndexes();
  } catch (e) {
    _dbg('json load error: $e');
  }
}

bool saveSettingsToExcel() {
  try {
    final f = File(settingsPath);
    f.parent.createSync(recursive: true);
    final excel = ex.Excel.createExcel();
    final order = terminals.entries.toList()..sort((a, b) => a.value.orderNum.compareTo(b.value.orderNum));
    excel.appendRow('Terminals', <ex.CellValue?>[
      ex.TextCellValue('Order'),
      ex.TextCellValue('TID'),
      ex.TextCellValue('Row'),
      ex.TextCellValue('Unit_Name'),
    ]);
    for (final e in order) {
      excel.appendRow('Terminals', <ex.CellValue?>[
        ex.IntCellValue(e.value.orderNum),
        ex.TextCellValue(e.key),
        ex.IntCellValue(e.value.excelRow),
        ex.TextCellValue(e.value.unitName),
      ]);
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
    try {
      excel.delete('Sheet1');
    } catch (_) {}
    try {
      File('${settingsPath}.json').writeAsStringSync(jsonEncode({
        'terminals': [
          for (final e in order)
            {'o': e.value.orderNum, 'tid': e.key, 'row': e.value.excelRow, 'name': e.value.unitName}
        ],
        'hints': [
          for (final h in nameHints) {'row': h.value, 'kws': h.key}
        ],
        'merchants': merchantCodeMap,
        'fixes': tidFix,
      }));
    } catch (_) {}
    final bytes = excel.save();
    if (bytes == null) return false;
    f.writeAsBytesSync(bytes);
    return true;
  } catch (e) {
    _dbg('settings save error: $e');
    return false;
  }
}

List<int>? _blankTemplateBytes() {
  try {
    final e = ex.Excel.createExcel();
    final name = e.getDefaultSheet() ?? 'Sheet1';
    e.appendRow(name, <ex.CellValue?>[
      null,
      ex.TextCellValue('الوحدة'),
      ex.TextCellValue('عدد الماكينات'),
      ex.TextCellValue('المبلغ'),
      ex.TextCellValue('كاش'),
      ex.TextCellValue('الإجمالي'),
    ]);
    return e.save();
  } catch (_) {
    return null;
  }
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
          try {
            tp.writeAsBytesSync(b);
          } catch (_) {}
        }
      }
    }
  } catch (e) {
    _dbg('default files error: $e');
  }
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
    addSheet('README', [
      'Terminals: Order = رقم الماكينة | TID = رقم الترمنل | Row = رقم الصف | Unit_Name = اسم الوحدة',
      'Name_Hints: كلمات مفتاحية لو ظهرت بالإيصال تحدد رقم الصف',
      'Merchant_Codes: كود التاجر -> رقم الصف',
      'TID_Fixes: أرقام تصحيح الـ OCR',
    ]);
    try {
      excel.delete('Sheet1');
    } catch (_) {}
    final bytes = excel.save();
    if (bytes == null) return false;
    f.writeAsBytesSync(bytes);
    return true;
  } catch (e) {
    _dbg('settings create error: $e');
    return false;
  }
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
          NavigationDestination(
            icon: _badged(homeBadge, Icons.receipt_long),
            label: tr('التقفيل'),
          ),
          NavigationDestination(
            icon: _badged(machBadge, Icons.folder_copy),
            label: tr('إيصالات الدفع'),
          ),
        ],
      ),
    );
  }
}

const String kAppName = 'A2R POS Public';
const String kAppAuthor = 'A2R Automated Systems';

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

  @override
  void didChangeLocales(List<Locale>? locales) {
    if (langCode == null) notifyListeners();
  }

  void setLang(String? code) {
    langCode = code;
    try {
      if (code == null) {
        _prefs?.remove('lang');
      } else {
        _prefs?.setString('lang', code);
      }
    } catch (_) {}
    notifyListeners();
  }

  void setMode(ThemeMode m) {
    themeMode = m;
    try {
      _prefs?.setString('theme_mode', m == ThemeMode.dark ? 'dark' : (m == ThemeMode.light ? 'light' : 'system'));
    } catch (_) {}
    notifyListeners();
  }

  void setSeed(int i) {
    seedIndex = i;
    try {
      _prefs?.setInt('seed', i);
    } catch (_) {}
    notifyListeners();
  }

  void toggleLang() {
    langCode = isAr ? 'en' : 'ar';
    try {
      _prefs?.setString('lang', langCode!);
    } catch (_) {}
    notifyListeners();
  }

  void setDark(bool dark) {
    themeMode = dark ? ThemeMode.dark : ThemeMode.light;
    try {
      _prefs?.setString('theme_mode', dark ? 'dark' : 'light');
    } catch (_) {}
    notifyListeners();
  }

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
  try {
    await Permission.notification.request();
  } catch (_) {}
  await askBatteryExempt();
  try {
    await _chan.invokeMethod('startKeepAlive', {'title': 'A2R POS', 'text': 'جاري الفحص الشامل (باتشات + ملفات)...'});
  } catch (_) {}
  try {
    await _chan.invokeMethod('keepScreenOn', {'on': true});
  } catch (_) {}
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
    try {
      await _chan.invokeMethod('keepScreenOn', {'on': false});
    } catch (_) {}
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
      builder: (c, sm, _) => sm.isEmpty
          ? const SizedBox.shrink()
          : Padding(padding: const EdgeInsets.only(bottom: 6), child: _statusBox(sm)),
    );

void _publishClosing(List<ReceiptItem> items) {
  final sums = <int, int>{};
  final bad = <int>{};
  for (final it in items) {
    if (!it.counts) continue;
    final n = it.orderNum ?? (it.tid != null ? terminals[it.tid]?.orderNum : null);
    if (n == null) continue;
    final t = it.txn;
    if (t == null) {
      bad.add(n);
      continue;
    }
    sums[n] = (sums[n] ?? 0) + t;
  }
  for (final n in bad) {
    sums.remove(n);
  }
  final cur = closingTxn.value;
  var same = cur.length == sums.length;
  if (same) {
    for (final e in sums.entries) {
      if (cur[e.key] != e.value) {
        same = false;
        break;
      }
    }
  }
  if (same) return;
  WidgetsBinding.instance.addPostFrameCallback((_) => closingTxn.value = sums);
}

Future<bool> confirmSend(BuildContext context, {int review = 0, int missing = 0, int pending = 0}) async {
  if (review == 0 && missing == 0 && pending == 0) return true;
  final ar = AppSettings.of(context).isAr;
  final lines = <String>[
    if (review > 0) (ar ? '• $review محتاج مراجعة' : '• $review need review'),
    if (missing > 0) (ar ? '• $missing ماكينة ناقصة' : '• $missing missing machine(s)'),
    if (pending > 0) (ar ? '• $pending لسه ما اتعدّش' : '• $pending not counted yet'),
  ];
  final r = await showDialog<bool>(
    context: context,
    builder: (ctx) => AlertDialog(
      title: w.Text(ar ? 'لسه فيه حاجات مفتوحة' : 'Some things are still open'),
      content: w.Text('${lines.join('\n')}\n\n${ar ? 'تبعت برضه؟' : 'Send anyway?'}'),
      actions: [
        TextButton(onPressed: () => Navigator.pop(ctx, false), child: w.Text(ar ? 'رجوع' : 'Back')),
        FilledButton(onPressed: () => Navigator.pop(ctx, true), child: w.Text(ar ? 'ابعت برضه' : 'Send anyway')),
      ],
    ),
  );
  return r == true;
}

Future<void> wipeEverything(BuildContext context) async {
  final ar = AppSettings.of(context).isAr;
  final m = ScaffoldMessenger.of(context);
  if (homeBusyG || machBusyG) {
    m.showSnackBar(SnackBar(
        content: w.Text(ar ? 'استنى لحد ما المعالجة تخلص أو وقّفها الأول.' : 'Wait for processing to finish or stop it first.')));
    return;
  }
  final ok = await showDialog<bool>(
    context: context,
    builder: (ctx) => AlertDialog(
      title: w.Text(ar ? 'مسح كل حاجة؟' : 'Wipe everything?'),
      content: w.Text(ar
          ? 'هيتمسح نهائي:\n• ملفات الإخراج (الإكسيل والـ PDF)\n• الصور المرتبة والنسخ المتسمية لإيصالات الدفع\n• نتايج الفحص والتصحيحات المحفوظة\n• اللوج\n• الكاش والملفات المؤقتة\n\nمش هيتمس: ملفات واتساب الأصلية، والإعدادات.'
          : 'Permanently wipe output files, scan caches, and logs?'),
      actions: [
        TextButton(onPressed: () => Navigator.pop(ctx, false), child: w.Text(ar ? 'إلغاء' : 'Cancel')),
        FilledButton(
          style: FilledButton.styleFrom(backgroundColor: Colors.red),
          onPressed: () => Navigator.pop(ctx, true),
          child: w.Text(ar ? 'امسح كل حاجة' : 'Wipe everything'),
        ),
      ],
    ),
  );
  if (ok != true) return;

  var files = 0;
  var bytes = 0;
  bool delFile(File f) {
    try {
      bytes += f.lengthSync();
      f.deleteSync();
      files++;
      return true;
    } catch (_) {
      return false;
    }
  }

  void wipeTree(String path) {
    try {
      final d = Directory(path);
      if (!d.existsSync()) return;
      for (final e in d.listSync(recursive: true, followLinks: false)) {
        if (e is File) delFile(e);
      }
      try {
        d.deleteSync(recursive: true);
      } catch (_) {}
    } catch (_) {}
  }

  void wipeFlat(String path, bool Function(File) only) {
    try {
      final d = Directory(path);
      if (!d.existsSync()) return;
      for (final e in d.listSync(followLinks: false)) {
        if (e is File && only(e)) delFile(e);
      }
    } catch (_) {}
  }

  bool appOut(File f) {
    final n = baseName(f.path);
    return n == 'review_report.txt' || n == 'ocr_debug.txt' || _isOutputFileName(n);
  }

  final protectedDirs = <String>{root, baseDir, '$root/Download', '$root/Documents', '$root/DCIM', '$root/Pictures'};
  for (final dir in {outputDir, defaultOutputDir}) {
    if (protectedDirs.contains(dir)) {
      wipeFlat(dir, appOut);
    } else if (dir == defaultOutputDir) {
      wipeTree(dir);
    } else {
      wipeFlat(dir, (f) => true);
    }
  }
  wipeTree(rawImagesDir);
  wipeTree('$baseDir/.scans');
  wipeTree(tmpDir);
  wipeTree(mTmpDir);
  wipeTree(defaultMDest);
  for (final p in [fixPath, mStatePath, '$baseDir/id_debug.txt']) {
    final f = File(p);
    if (f.existsSync()) delFile(f);
  }
  try {
    final cfg = _readCfg();
    cfg.remove('doneM');
    _writeCfg(cfg);
  } catch (_) {}
  imageCache.clear();
  imageCache.clearLiveImages();
  closingTxn.value = <int, int>{};
  wipeSignal.value++;
  final mb = (bytes / 1048576).toStringAsFixed(1);
  m.showSnackBar(SnackBar(
      content: w.Text(ar ? 'تم المسح: $files ملف ($mb ميجا) 🧹' : 'Wiped: $files files ($mb MB) 🧹')));
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
    appBarTheme: AppBarTheme(
      backgroundColor: cs.surface,
      foregroundColor: cs.onSurface,
      surfaceTintColor: Colors.transparent,
      scrolledUnderElevation: 1,
    ),
    cardTheme: CardThemeData(
      elevation: 0,
      color: cs.surfaceContainerLow,
      shape: RoundedRectangleBorder(
        borderRadius: BorderRadius.circular(14),
        side: BorderSide(color: cs.outlineVariant),
      ),
    ),
    navigationBarTheme: NavigationBarThemeData(
      backgroundColor: cs.surfaceContainer,
      indicatorColor: cs.secondaryContainer,
      labelBehavior: NavigationDestinationLabelBehavior.alwaysShow,
    ),
    inputDecorationTheme: InputDecorationTheme(
      border: OutlineInputBorder(borderRadius: BorderRadius.circular(10)),
    ),
  );
}

class _LangScope extends InheritedNotifier<AppSettings> {
  const _LangScope({Key? key, required AppSettings notifier, required Widget child})
      : super(key: key, notifier: notifier, child: child);
}

class Text extends StatelessWidget {
  const Text(
    this.data, {
    super.key,
    this.style,
    this.textAlign,
    this.textDirection,
    this.softWrap,
    this.overflow,
    this.maxLines,
  });

  final String data;
  final TextStyle? style;
  final TextAlign? textAlign;
  final TextDirection? textDirection;
  final bool? softWrap;
  final TextOverflow? overflow;
  final int? maxLines;

  @override
  Widget build(BuildContext context) {
    AppSettings.of(context);
    return w.Text(
      tr(data),
      style: style,
      textAlign: textAlign,
      textDirection: textDirection,
      softWrap: softWrap,
      overflow: overflow,
      maxLines: maxLines,
    );
  }
}

const String _arL = '\u0621-\u063A\u0641-\u064A\u064B-\u0652\u0671-\u06D3\u0640';
final RegExp _hasArabic = RegExp('[$_arL]');
final Map<String, String> _trCache = {};

String tr(String s) {
  if (s.isEmpty || appSettings.isAr || !_hasArabic.hasMatch(s)) return s;
  return s; // يتم ترجمتها عبر القاموس إن وجدت
}

bool _isDark(BuildContext c) => Theme.of(c).brightness == Brightness.dark;

Color _tintBg(BuildContext c, MaterialColor m, {bool soft = false, bool strong = false}) {
  if (_isDark(c)) return m.shade900.withAlpha(strong ? 210 : 140);
  return strong ? m.shade300 : (soft ? m.shade50 : m.shade100);
}

Color _tintBorder(BuildContext c, MaterialColor m) => _isDark(c) ? m.shade300 : m.shade800;
Color _tintFg(BuildContext c, MaterialColor m) => _isDark(c) ? m.shade200 : m.shade900;

class AppTitle extends StatelessWidget {
  const AppTitle({super.key});

  @override
  Widget build(BuildContext context) {
    final cs = Theme.of(context).colorScheme;
    return Column(
      mainAxisSize: MainAxisSize.min,
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        const w.Text(
          kAppName,
          textDirection: TextDirection.ltr,
          style: TextStyle(fontSize: 17, fontWeight: FontWeight.bold, height: 1.1),
        ),
        w.Text(
          kAppAuthor,
          textDirection: TextDirection.ltr,
          style: TextStyle(fontSize: 10, height: 1.2, color: cs.onSurfaceVariant),
        ),
      ],
    );
  }
}

class HeaderButtons extends StatelessWidget {
  const HeaderButtons({super.key});

  @override
  Widget build(BuildContext context) {
    if (MediaQuery.sizeOf(context).width < 400) return const SizedBox.shrink();
    final ar = AppSettings.of(context).isAr;
    return IconButton(
      tooltip: ar ? 'المظهر واللغة' : 'Theme & language',
      icon: Icon(_isDark(context) ? Icons.dark_mode : Icons.palette_outlined),
      onPressed: () => showAppearanceSheet(context),
    );
  }
}

void showAppearanceSheet(BuildContext context) {
  showModalBottomSheet<void>(
    context: context,
    showDragHandle: true,
    isScrollControlled: true,
    builder: (_) => const _AppearanceSheet(),
  );
}

class _AppearanceSheet extends StatelessWidget {
  const _AppearanceSheet();

  @override
  Widget build(BuildContext context) {
    final s = AppSettings.of(context);
    final ar = s.isAr;
    final cs = Theme.of(context).colorScheme;
    Widget label(String t) => Padding(
          padding: const EdgeInsets.fromLTRB(4, 14, 4, 8),
          child: w.Text(t, style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 13)),
        );
    final order = [kDefaultSeed, for (var i = 0; i < kSeeds.length; i++) if (i != kDefaultSeed) i];
    return SafeArea(
      child: SingleChildScrollView(
        padding: const EdgeInsets.fromLTRB(16, 0, 16, 16),
        child: Column(
          mainAxisSize: MainAxisSize.min,
          crossAxisAlignment: CrossAxisAlignment.stretch,
          children: [
            label(ar ? 'اللغة' : 'Language'),
            SegmentedButton<String>(
              showSelectedIcon: false,
              segments: [
                ButtonSegment<String>(value: 'sys', label: w.Text(ar ? 'تلقائي' : 'Auto')),
                const ButtonSegment<String>(value: 'ar', label: w.Text('العربية')),
                const ButtonSegment<String>(value: 'en', label: w.Text('English')),
              ],
              selected: {s.langCode ?? 'sys'},
              onSelectionChanged: (v) => s.setLang(v.first == 'sys' ? null : v.first),
            ),
            label(ar ? 'المظهر' : 'Theme'),
            SegmentedButton<ThemeMode>(
              showSelectedIcon: false,
              segments: [
                ButtonSegment<ThemeMode>(
                  value: ThemeMode.system,
                  icon: const Icon(Icons.brightness_auto, size: 18),
                  label: w.Text(ar ? 'تلقائي' : 'Auto'),
                ),
                ButtonSegment<ThemeMode>(
                  value: ThemeMode.light,
                  icon: const Icon(Icons.light_mode, size: 18),
                  label: w.Text(ar ? 'فاتح' : 'Light'),
                ),
                ButtonSegment<ThemeMode>(
                  value: ThemeMode.dark,
                  icon: const Icon(Icons.dark_mode, size: 18),
                  label: w.Text(ar ? 'داكن' : 'Dark'),
                ),
              ],
              selected: {s.themeMode},
              onSelectionChanged: (v) => s.setMode(v.first),
            ),
            label(ar ? 'لون التطبيق' : 'Accent color'),
            Wrap(
              spacing: 12,
              runSpacing: 12,
              children: [
                for (final i in order)
                  Tooltip(
                    message: ar ? kSeedNamesAr[i] : kSeedNamesEn[i],
                    child: InkWell(
                      customBorder: const CircleBorder(),
                      onTap: () => s.setSeed(i),
                      child: Container(
                        width: 40,
                        height: 40,
                        decoration: BoxDecoration(
                          color: kSeeds[i],
                          shape: BoxShape.circle,
                          border: Border.all(
                            color: s.seedIndex == i ? cs.onSurface : Colors.transparent,
                            width: 3,
                          ),
                        ),
                        child: s.seedIndex == i
                            ? const Icon(Icons.check, color: Colors.white, size: 20)
                            : null,
                      ),
                    ),
                  ),
              ],
            ),
          ],
        ),
      ),
    );
  }
}

// ───────────────────── نموذج الماكينات والوحدات الديناميكي ─────────────────────
// خالٍ تماماً من أي أكواد أو أسماء مدن أو وحدات مسبقة
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

// ───────────────────── استخراج البيانات والـ OCR المتقدم ─────────────────────
final List<RegExp> _markers = [
  RegExp(r'TERMINAL', caseSensitive: false),
  RegExp(r'GRAND|TOTALS?', caseSensitive: false),
  RegExp(r'NATIONAL\s*BANK|\bNBE\b', caseSensitive: false),
  RegExp(r'BATCH', caseSensitive: false),
  RegExp(r'MERCHANT', caseSensitive: false),
  RegExp(r'SUCCESSF|CLOSED', caseSensitive: false),
  RegExp(r'\bVISA\b|MASTER|MEZZA|MEEZA|MEEZ', caseSensitive: false),
  RegExp(r'THANK\s*YOU', caseSensitive: false),
  RegExp(r'NET\s*AMOUNT', caseSensitive: false),
  RegExp(r'\bSALE\b', caseSensitive: false),
];

int receiptScore(String text) => _markers.where((r) => r.hasMatch(text)).length;

class OcrLine {
  final String text;
  final double left, right, top, bottom;
  OcrLine(this.text, this.left, this.right, this.top, this.bottom);
  double get cy => (top + bottom) / 2;
  double get h => bottom - top;
}

List<OcrLine> toLines(RecognizedText res) {
  final out = <OcrLine>[];
  for (final b in res.blocks) {
    for (final l in b.lines) {
      final t = l.text.trim();
      if (t.isEmpty) continue;
      final r = l.boundingBox;
      out.add(OcrLine(t, r.left, r.right, r.top, r.bottom));
    }
  }
  return out;
}

String rowsText(List<OcrLine> lines) {
  final ls = [...lines]..sort((a, b) => a.cy.compareTo(b.cy));
  final rows = <List<OcrLine>>[];
  for (final l in ls) {
    if (rows.isNotEmpty) {
      final last = rows.last;
      final cy = last.map((e) => e.cy).reduce((a, b) => a + b) / last.length;
      final h = last.map((e) => e.h).reduce((a, b) => a + b) / last.length;
      if ((l.cy - cy).abs() < h * 0.6) {
        last.add(l);
        continue;
      }
    }
    rows.add([l]);
  }
  final sb = StringBuffer();
  for (final r in rows) {
    r.sort((a, b) => a.left.compareTo(b.left));
    sb.writeln(r.map((e) => e.text).join('  '));
  }
  return sb.toString();
}

String readingText(List<OcrLine> lines) => lines.map((e) => e.text).join('\n');

List<List<OcrLine>> splitColumns(List<OcrLine> lines) {
  if (lines.length < 8) return [lines];
  final minL = lines.map((e) => e.left).reduce(math.min);
  final maxR = lines.map((e) => e.right).reduce(math.max);
  final w = maxR - minL;
  if (w <= 0) return [lines];
  final sorted = [...lines]..sort((a, b) => a.left.compareTo(b.left));
  for (final tolFrac in const [0.03, 0.02, 0.01]) {
    final groups = <List<OcrLine>>[];
    var curRight = -1e9;
    for (final l in sorted) {
      if (groups.isEmpty || l.left - curRight > w * tolFrac) {
        groups.add([l]);
        curRight = l.right;
      } else {
        groups.last.add(l);
        curRight = math.max(curRight, l.right);
      }
    }
    final big = groups.where((g) {
      final gl = g.map((e) => e.left).reduce(math.min);
      final gr = g.map((e) => e.right).reduce(math.max);
      return gr - gl >= w * 0.12 && g.length >= 6;
    }).toList();
    if (big.length >= 2 &&
        big.length <= 4 &&
        big.every((g) => receiptScore(readingText(g)) >= 3)) {
      return big;
    }
  }
  return [lines];
}

final RegExp _amountRe =
    RegExp(r"(?<![\d,.])(\d{1,3}(?:,\d{3})*|\d+)[., ](\d{2})(?!\d)");
final RegExp _brandRe = RegExp(
    r"\b(VISA|MASTER(?:CARD)?|MEZZA|MEEZA|MEZA|MEZZ4|MEZZ)\b",
    caseSensitive: false);

List<double> amountsIn(String rawText) {
  final text = rawText.replaceAllMapped(
      RegExp(r'\bB(\d{2,3}[.,]\d{2})\b'), (m) => '8${m.group(1)}');
  final out = <double>[];
  for (final m in _amountRe.allMatches(text)) {
    final v = double.tryParse("${m.group(1)!.replaceAll(',', '')}.${m.group(2)!}");
    if (v != null && v >= 0.5 && v <= 5000000) out.add(round2(v));
  }
  return out;
}

MapEntry<T, int>? mostCommon<T>(List<T> vals) {
  if (vals.isEmpty) return null;
  final counts = <T, int>{};
  for (final v in vals) {
    counts[v] = (counts[v] ?? 0) + 1;
  }
  MapEntry<T, int>? best;
  for (final e in counts.entries) {
    if (best == null || e.value > best.value) best = e;
  }
  return best;
}

class Parsed {
  final List<double> grand;
  final double? bsum;
  final List<double> bvals;
  Parsed(this.grand, this.bsum, this.bvals);
}

Parsed parseReceipt(String txt) {
  final t = txt
      .split(RegExp(r"THANK\s*YOU|Hotline|App\s*Version", caseSensitive: false))
      .first;
  final gm = RegExp(r"GRAND|TOTALS", caseSensitive: false).firstMatch(t);
  var pre = gm != null ? t.substring(0, gm.start) : t;
  final post = gm != null ? t.substring(gm.start) : "";

  final bms = _brandRe.allMatches(pre).toList();
  final sections = <String>[];
  if (bms.isNotEmpty) {
    for (var i = 0; i < bms.length; i++) {
      final end = i + 1 < bms.length ? bms[i + 1].start : pre.length;
      sections.add(pre.substring(bms[i].end, end));
    }
  } else {
    final sm =
        RegExp(r"SUCCESSFULL?Y|CLOSED", caseSensitive: false).firstMatch(pre);
    if (sm != null) pre = pre.substring(sm.end);
    sections.add(pre);
  }

  final bvals = <double>[];
  for (final sec in sections) {
    final mc = mostCommon(amountsIn(sec));
    if (mc != null) bvals.add(mc.key);
  }
  final grand = amountsIn(post);
  final bsum = bvals.isNotEmpty ? round2(bvals.reduce((a, b) => a + b)) : null;
  return Parsed(grand, bsum, bvals);
}

List<double> candidateAmounts(List<String> texts) {
  final grandScore = <double, int>{};
  final otherScore = <double, double>{};
  for (final t in texts) {
    final pr = parseReceipt(t);
    for (final g in pr.grand) {
      if (g >= 0.5) grandScore[g] = (grandScore[g] ?? 0) + 1;
    }
    final b = pr.bsum;
    if (b != null && b > 0) otherScore[round2(b)] = (otherScore[round2(b)] ?? 0) + 3;
    for (final v in pr.bvals) {
      if (v >= 0.5) otherScore[v] = (otherScore[v] ?? 0) + 1;
    }
  }
  if (grandScore.isNotEmpty) {
    final ks = grandScore.keys.toList()..sort((a, b) => grandScore[b]!.compareTo(grandScore[a]!));
    return ks.take(3).toList();
  }
  final ks = otherScore.keys.toList()..sort((a, b) => otherScore[b]!.compareTo(otherScore[a]!));
  return ks.take(4).toList();
}

List<double> mergeCands(List<double> a, List<double> b) {
  bool has(List<double> l, double v) => l.any((x) => (x - v).abs() < 0.005);
  final out = <double>[];
  for (final v in a) {
    if (has(b, v)) out.add(v);
  }
  for (final v in [...a, ...b]) {
    if (!has(out, v)) out.add(v);
  }
  return out.take(4).toList();
}

class Decision {
  final double val;
  final bool ok;
  Decision(this.val, this.ok);
}

Decision decide(List<Parsed> results) {
  for (final r in results) {
    final b = r.bsum;
    if (b != null && b > 0 && r.grand.any((g) => (g - b).abs() < 0.005)) {
      return Decision(b, true);
    }
  }
  for (final r in results) {
    final b = r.bsum;
    if (b != null && b > 0 && r.bvals.length >= 2) {
      for (final g in r.grand) {
        final sa = g.toStringAsFixed(2), sb = b.toStringAsFixed(2);
        if (sa.length == sb.length) {
          var diff = 0;
          for (var i = 0; i < sa.length; i++) {
            if (sa[i] != sb[i]) diff++;
          }
          if (diff == 1) return Decision(b, true);
        }
      }
    }
  }
  for (final r in results) {
    final c = mostCommon(r.grand);
    if (c != null && c.value >= 2 && r.bsum == null) return Decision(c.key, false);
  }
  final pool = <double>[];
  for (final r in results) {
    pool.addAll(r.grand);
    final b = r.bsum;
    if (b != null && b > 0) pool.add(b);
  }
  final c = mostCommon(pool);
  if (c != null) return Decision(c.key, false);
  return Decision(0.0, false);
}

class PieceResult {
  final List<String> texts;
  final double val;
  final bool ok;
  PieceResult(this.texts, this.val, this.ok);
}

String classifyReceipt(String text) {
  final up = text.toUpperCase();
  final payRes = <RegExp>[
    RegExp(r'R{1,2}N\W{0,3}\d{9,}'),
    RegExp(r'\d{4,6}\W{0,2}\*{3,}\W{0,2}\d{4}'),
    RegExp(r'APPROVED'),
    RegExp(r'AUTH\W{0,3}\d{6}'),
    RegExp(r'I\s*AUTHORI[SZ]E'),
    RegExp(r'MERCHANT\s*COPY'),
    RegExp(r'CLESS|CHIP\s*ONLINE'),
  ];
  final batchRes = <RegExp>[
    RegExp(r'BATCH\s*CLOS'),
    RegExp(r'SUCCESSF'),
    RegExp(r'GRAND\s*TOTAL'),
    RegExp(r'TOTAL\s*NET'),
    RegExp(r'THANK\s*YOU'),
    RegExp(r'NBE\s*NETWORK'),
    RegExp(r'HOTLINE'),
  ];
  final pay = payRes.where((r) => r.hasMatch(up)).length;
  final batch = batchRes.where((r) => r.hasMatch(up)).length;
  if (pay >= 2 && batch < 2) return 'payment';
  if (batch >= 1) return 'batch';
  return 'other';
}

double? sumClosure(List<String> texts) {
  final cnt = <double, int>{};
  for (final t in texts) {
    final tt = t
        .split(RegExp(r'THANK\s*YOU|Hotline|App\s*Version', caseSensitive: false))
        .first;
    final local = <double, int>{};
    for (final a in amountsIn(tt)) {
      local[a] = (local[a] ?? 0) + 1;
    }
    local.forEach((k, v) {
      if (v > (cnt[k] ?? 0)) cnt[k] = v;
    });
  }
  final c = cnt.entries.where((e) => e.value >= 2).map((e) => e.key).toList()..sort();
  if (c.length < 3 || c.length > 14) return null;
  for (var gi = c.length - 1; gi >= 2; gi--) {
    final g = c[gi];
    final others = c.sublist(0, gi);
    final n = others.length;
    for (var mask = 1; mask < (1 << n); mask++) {
      var bits = 0, sum = 0.0;
      for (var i = 0; i < n; i++) {
        if (mask & (1 << i) != 0) {
          bits++;
          sum += others[i];
        }
      }
      if (bits >= 2 && (sum - g).abs() < 0.005) return g;
    }
  }
  return null;
}

PieceResult ocrPiece(List<OcrLine> lines) {
  final texts = <String>[rowsText(lines), readingText(lines)];
  final sc = sumClosure(texts);
  if (sc != null) return PieceResult(texts, sc, true);
  final d = decide(texts.map(parseReceipt).toList());
  return PieceResult(texts, d.val, d.ok);
}

String? batchNo(List<String> texts) {
  final reg = RegExp(r"BATCH\s*NO\D{0,6}(\d{4,7})", caseSensitive: false);
  final all = <String>[];
  for (final t in texts) {
    for (final m in reg.allMatches(t)) {
      all.add(m.group(1)!);
    }
  }
  return mostCommon(all)?.key;
}

String? _tidLabeledRaw(String t) {
  final m = RegExp(r"TERMINAL\s*(?:ID)?\D{0,15}?(\d[\d\s/]{6,13}\d)",
          caseSensitive: false)
      .firstMatch(t);
  if (m == null) return null;
  final raw = m.group(1)!.replaceAll('/', '7').replaceAll(RegExp(r'\D'), '');
  if (raw.length < 8 || raw.length > 9 || raw.startsWith('0011')) return null;
  final d = raw.substring(raw.length - 8);
  return tidFix[d] ?? d;
}

String? _tidLabeled(String t) {
  final d = _tidLabeledRaw(t);
  return d != null && terminals.containsKey(d) ? d : null;
}

int _editDistance(String a, String b) {
  if (a == b) return 0;
  if (a.isEmpty) return b.length;
  if (b.isEmpty) return a.length;
  var prev = List<int>.generate(b.length + 1, (i) => i);
  for (var i = 1; i <= a.length; i++) {
    final cur = List<int>.filled(b.length + 1, 0);
    cur[0] = i;
    for (var j = 1; j <= b.length; j++) {
      final cost = a[i - 1] == b[j - 1] ? 0 : 1;
      cur[j] = math.min(math.min(cur[j - 1] + 1, prev[j] + 1), prev[j - 1] + cost);
    }
    prev = cur;
  }
  return prev[b.length];
}

String _normUnit(String s) => s.toUpperCase().replaceAll(RegExp(r'[^A-Z0-9\u0600-\u06FF]'), '');

String? _cleanUnitName(String raw) {
  final s = raw
      .replaceAll(RegExp(r'[^A-Za-z0-9\u0600-\u06FF \-]'), ' ')
      .replaceAll(RegExp(r'\s+'), ' ')
      .trim();
  final letters = RegExp(r'[A-Za-z\u0600-\u06FF]').allMatches(s).length;
  final digits = RegExp(r'\d').allMatches(s).length;
  if (letters < 4 || digits > letters) return null;
  return s.length > 40 ? s.substring(0, 40).trim() : s;
}

String? extractUnitName(List<String> texts) {
  final label = RegExp(r'MERCHANT\s*NAME\s*[:\-]?\s*(.+)', caseSensitive: false);
  for (final t in texts) {
    for (final l in t.split('\n')) {
      final m = label.firstMatch(l.trim());
      if (m != null) {
        final n = _cleanUnitName(m.group(1)!);
        if (n != null) return n;
      }
    }
  }
  final ecRe = RegExp(r'\bE\s?C\s+([A-Z0-9][A-Z0-9 \-]{2,40}?)\s+TR[FE]\b');
  final ecFound = <String>[];
  for (final t in texts) {
    for (final l in t.split('\n')) {
      final m = ecRe.firstMatch(l.toUpperCase());
      if (m != null) {
        final n = _cleanUnitName('EC ${m.group(1)!.trim()} TRF');
        if (n != null) ecFound.add(n);
      }
    }
  }
  final ecBest = mostCommon(ecFound)?.key;
  if (ecBest != null) return ecBest;
  final skip = RegExp(
      r'NATIONAL|BANK|\bNBE\b|EGYPT|RECEIPT|CUSTOMER|COPY|MERCHANT|TERMINAL|BATCH|DATE|TIME|SALE|TOTAL|AMOUNT|VISA|MASTER|MEEZ|THANK|SUCCESS|CLOSED|CARD|APPROVAL|\bRRN\b|TRACE|VERSION|GRAND|REPORT|SETTLEMENT|\bST\b|STREET|ALEX|CAIRO|\b02F\b|INSIDE|NEXT\s*TO|\bFRON',
      caseSensitive: false);
  final bankLine = RegExp(r'BANK\s*OF\s*EGYPT|NATIONAL\s*BANK|^\s*NBE\s*$', caseSensitive: false);
  for (final t in texts) {
    final ls = t.split('\n').map((e) => e.trim()).where((e) => e.isNotEmpty).toList();
    for (var i = 0; i < ls.length; i++) {
      if (!bankLine.hasMatch(ls[i])) continue;
      for (var j = i + 1; j < ls.length && j <= i + 3; j++) {
        if (skip.hasMatch(ls[j])) continue;
        final n = _cleanUnitName(ls[j]);
        if (n != null) return n;
      }
      break;
    }
  }
  for (final t in texts) {
    var seen = 0;
    for (final raw in t.split('\n')) {
      final l = raw.trim();
      if (l.isEmpty) continue;
      if (++seen > 8) break;
      if (skip.hasMatch(l)) continue;
      final n = _cleanUnitName(l);
      if (n != null) return n;
    }
  }
  return null;
}

int _rowForUnitName(String name) {
  final k = _normUnit(name);
  for (final e in rowNames.entries) {
    final o = _normUnit(e.value);
    if (o == k) return e.key;
    final mx = math.max(o.length, k.length);
    if (mx >= 6 && _editDistance(o, k) <= (mx * 0.15).floor()) return e.key;
  }
  var mxRow = 1;
  for (final r in rowNames.keys) {
    if (r > mxRow) mxRow = r;
  }
  return mxRow + 1;
}

String? extractMerchantId(List<String> texts) {
  final found = <String>[];
  final re = RegExp(r"MERCHANT\s*(?:ID|10|NO)\D{0,15}?(\d[\d\s/]{6,16}\d)", caseSensitive: false);
  for (final t in texts) {
    for (final m in re.allMatches(t)) {
      final d = m.group(1)!.replaceAll('/', '7').replaceAll(RegExp(r'\D'), '');
      if (d.length >= 8 && d.length <= 16) found.add(d);
    }
  }
  return mostCommon(found)?.key;
}

int? _rowForMerchant(String mid) {
  final exact = merchantCodeMap[mid];
  if (exact != null) return exact;
  for (final e in merchantCodeMap.entries) {
    final k = e.key;
    if (k.length >= 10 && mid.length > k.length && mid.contains(k)) return e.value;
  }
  return null;
}

void learnMerchant(String tid, List<String> texts) {
  final t = terminals[tid];
  if (t == null) return;
  final mid = extractMerchantId(texts);
  if (mid == null || _rowForMerchant(mid) != null) return;
  merchantCodeMap[mid] = t.excelRow;
  saveSettingsToExcel();
}

String? discoverTerminal(List<String> texts) {
  final counts = <String, int>{};
  final midNow = extractMerchantId(texts);
  bool inMerchant(String d) => midNow != null && midNow.contains(d);
  for (final t in texts) {
    final d = _tidLabeledRaw(t);
    if (d != null && !inMerchant(d)) counts[d] = (counts[d] ?? 0) + 1;
  }
  if (counts.isEmpty) {
    for (final t in texts) {
      for (final m in RegExp(r'(?<!\d)(50\d{6})(?!\d)').allMatches(t)) {
        final d = tidFix[m.group(1)!] ?? m.group(1)!;
        if (!inMerchant(d)) counts[d] = (counts[d] ?? 0) + 1;
      }
    }
  }
  if (counts.isEmpty) return null;
  final tid = counts.entries.reduce((a, b) => a.value >= b.value ? a : b).key;
  if (terminals.containsKey(tid)) return tid;
  final name = extractUnitName(texts) ?? 'Unit $tid';
  final mid = extractMerchantId(texts);
  final byMerchant = mid == null ? null : _rowForMerchant(mid);
  final row = byMerchant ?? _rowForUnitName(name);
  var order = 0;
  for (final t in terminals.values) {
    if (t.orderNum > order) order = t.orderNum;
  }
  terminals[tid] = TerminalInfo(order + 1, row, rowNames[row] ?? name);
  if (mid != null && byMerchant == null) merchantCodeMap[mid] = row;
  rebuildTerminalIndexes();
  saveSettingsToExcel();
  return tid;
}

String? _tidLoose(String t) {
  for (final tid in terminals.keys) {
    if (t.contains(tid)) return tid;
  }
  for (final m in RegExp(r"\b(\d{8})\b").allMatches(t)) {
    var c = m.group(1)!;
    c = tidFix[c] ?? c;
    if (terminals.containsKey(c)) return c;
  }
  return null;
}

String? findTerminalIn(List<String> texts) {
  for (final t in texts) {
    final r = _tidLabeled(t);
    if (r != null) return r;
  }
  for (final t in texts) {
    final r = _tidLoose(t);
    if (r != null) return r;
  }
  return null;
}

String? findTerminalFuzzyIn(List<String> texts, int? targetRow) {
  final pool = terminals.entries
      .where((e) => targetRow == null || e.value.excelRow == targetRow)
      .map((e) => e.key)
      .toList();
  final dist = <String, int>{};
  for (final t in texts) {
    for (final run in RegExp(r"\d[\d /]{6,14}\d").allMatches(t)) {
      final digits = run.group(0)!.replaceAll('/', '7').replaceAll(RegExp(r'\D'), '');
      for (var i = 0; i + 8 <= digits.length; i++) {
        final cand = digits.substring(i, i + 8);
        for (final tid in pool) {
          var d = 0;
          for (var k = 0; k < 8; k++) {
            if (cand[k] != tid[k]) d++;
          }
          if (d < (dist[tid] ?? 99)) dist[tid] = d;
        }
      }
    }
  }
  if (dist.isEmpty) return null;
  final best = dist.values.reduce(math.min);
  final winners = dist.entries.where((e) => e.value == best).map((e) => e.key).toList();
  if (best <= 1 && winners.length == 1) return winners.first;
  return null;
}

int? findUnitExtraIn(List<String> texts) {
  for (final t in texts) {
    final m = RegExp(r"MERCHANT\s*(?:ID|10)\D{0,15}?(\d[\d\s/]{8,13}\d)", caseSensitive: false)
        .firstMatch(t);
    final cands = <String>[
      if (m != null) m.group(1)!.replaceAll('/', '7').replaceAll(RegExp(r'\D'), ''),
      t.replaceAll(RegExp(r'\D'), ''),
    ];
    for (final d in cands) {
      for (final e in merchantCodeMap.entries) {
        if (e.key.length >= 6 && d.contains(e.key)) return e.value;
      }
    }
  }
  for (final t in texts) {
    final up = t.toUpperCase();
    for (final hint in nameHints) {
      if (hint.key.any((w) => up.contains(w))) return hint.value;
    }
  }
  return null;
}

String? findDateIn(List<String> texts) {
  final reg = RegExp(r"DATE\s*:?\s*(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})",
      caseSensitive: false);
  for (final t in texts) {
    final m = reg.firstMatch(t);
    if (m != null) {
      final mo = int.tryParse(m.group(2)!) ?? 0;
      final d = int.tryParse(m.group(3)!) ?? 0;
      return "$d.$mo";
    }
  }
  return null;
}

int? _txnNum(String s) {
  final t = s.trim().replaceAll(RegExp(r'[Oo]'), '0').replaceAll(RegExp(r'[Il|]'), '1');
  return int.tryParse(t);
}

int? batchTxnOf(List<String> texts) {
  final grandRe = RegExp(r'GRAND|TOTALS', caseSensitive: false);
  final paren = RegExp(r'\(\s*([0-9OoIl|]{1,5})\s*\)');
  final grands = <int>[];
  final sums = <int>[];
  for (final raw in texts) {
    final gm = grandRe.firstMatch(raw);
    if (gm != null) {
      final m = paren.firstMatch(raw.substring(gm.start));
      final v = m == null ? null : _txnNum(m.group(1)!);
      if (v != null && v > 0 && v < 100000) grands.add(v);
    }
    final pre = gm != null ? raw.substring(0, gm.start) : raw;
    var sum = 0, n = 0;
    for (final m in RegExp(r'TOTAL(?!\s*NET)[^()\n]{0,12}\(\s*([0-9OoIl|]{1,5})\s*\)', caseSensitive: false)
        .allMatches(pre)) {
      final v = _txnNum(m.group(1)!);
      if (v != null) {
        sum += v;
        n++;
      }
    }
    if (n > 0 && sum > 0) sums.add(sum);
  }
  final g = mostCommon(grands)?.key;
  final s = mostCommon(sums)?.key;
  if (g != null && s != null && g != s) {
    if (g > 0 && g < 1000) return g;
  }
  return g ?? s;
}

class Analysis {
  String? tid;
  String how;
  int? row;
  String? unitName;
  double val;
  bool ok;
  List<double> batches;
  String? date;
  String debug = '';
  int? txn;
  List<double> cands = <double>[];
  Analysis(this.tid, this.how, this.row, this.unitName, this.val, this.ok,
      this.batches, this.date);
}

Analysis analyzeLines(List<OcrLine> all) {
  final pieces = splitColumns(all);
  final multi = pieces.length > 1;
  final allTexts = <String>[];
  final batches = <double>[];
  final txns = <int?>[];
  var allOk = true;
  final seen = <String>{};

  final dbg = StringBuffer('pieces=${pieces.length}\n');
  for (final pc in pieces) {
    final r = ocrPiece(pc);
    final pr = parseReceipt(r.texts[0]);
    dbg.writeln('--- receipt ---\n${r.texts[0]}>> cards=${pr.bvals} sum=${pr.bsum} '
        'grand=${pr.grand} => ${r.val} ok=${r.ok}');
    allTexts.addAll(r.texts);
    if (r.val <= 0) {
      allOk = false;
      continue;
    }
    final b = batchNo(r.texts);
    final key = '$b|${r.val}';
    if (b != null && seen.contains(key)) continue;
    seen.add(key);
    batches.add(r.val);
    txns.add(batchTxnOf(r.texts));
    allOk = allOk && r.ok;
  }
  var total = round2(batches.fold(0.0, (a, b) => a + b));
  var ok = allOk && batches.isNotEmpty;

  if (multi && !ok) {
    final r = ocrPiece(all);
    allTexts.addAll(r.texts);
    if (r.val > 0 && (r.ok || batches.isEmpty)) {
      total = r.val;
      ok = r.ok;
      batches
        ..clear()
        ..add(r.val);
      txns
        ..clear()
        ..add(batchTxnOf(r.texts));
    }
  }

  final date = findDateIn(allTexts);
  var tid = findTerminalIn(allTexts) ?? discoverTerminal(allTexts);
  if (tid != null) learnMerchant(tid, allTexts);
  var how = 'terminal';
  int? row;
  String? name;
  if (tid == null) {
    final u0 = findUnitExtraIn(allTexts);
    tid = findTerminalFuzzyIn(allTexts, u0);
    how = 'fuzzy';
  }
  if (tid != null) {
    tid = tidAlias[tid] ?? tid;
    final info = terminals[tid]!;
    row = info.excelRow;
    name = info.unitName;
  } else {
    how = 'none';
    final u = findUnitExtraIn(allTexts);
    if (u != null) {
      row = u;
      name = rowNames[u];
      how = 'name/merchant';
    }
  }
  final result = Analysis(tid, how, row, name, total, ok, batches, date);
  result.debug = dbg.toString();
  result.cands = candidateAmounts(allTexts);
  result.txn = (txns.isNotEmpty && txns.every((e) => e != null)) ? txns.fold<int>(0, (a, e) => a + e!) : null;
  return result;
}

double round2(double val) => double.parse(val.toStringAsFixed(2));

double? mergeSecondRead(Analysis a, Analysis b) {
  a.cands = mergeCands(a.cands, b.cands);
  if (a.tid == null && b.tid != null) {
    a.tid = b.tid;
    a.row = b.row;
    a.unitName = b.unitName;
    a.how = b.how;
  } else if (a.row == null && b.row != null) {
    a.row = b.row;
    a.unitName = b.unitName;
    a.how = b.how;
  }
  a.date ??= b.date;
  a.txn ??= b.txn;
  if (b.val <= 0) return null;
  if (a.val <= 0) {
    a.val = b.val;
    a.ok = b.ok;
    a.batches = [...b.batches];
    a.txn = b.txn ?? a.txn;
    return null;
  }
  if ((a.val - b.val).abs() < 0.005) return null;
  if (!a.ok && b.ok) {
    a.val = b.val;
    a.ok = true;
    a.batches = [...b.batches];
    a.txn = b.txn ?? a.txn;
    return null;
  }
  if (a.ok && !b.ok) return null;
  a.ok = false;
  return b.val;
}

Future<List<File>> extractPdfPages(File pdf) async {
  final out = <File>[];
  try {
    final bytes = await pdf.readAsBytes();
    final dir = Directory('$tmpDir/pdf')..createSync(recursive: true);
    var i = 0;
    await for (final page in Printing.raster(bytes, dpi: 220)) {
      i++;
      final png = await page.toPng();
      final f = File('${dir.path}/pdf${bytes.length}_p$i.png');
      await f.writeAsBytes(png);
      out.add(f);
    }
  } catch (_) {}
  return out;
}

// ───────────────────── مسارات المجلدات والتهيئة ─────────────────────
const String root = '/storage/emulated/0';
final String baseDir = '$root/A2R-POS-Public';
final String rawImagesDir = '$baseDir/raw_images';
String outputDir = '$baseDir';
final String defaultOutputDir = '$baseDir';
final String tmpDir = '$baseDir/.tmp';
final String mTmpDir = '$baseDir/.tmp_m';
void _ensureMTmp() {
  try {
    Directory(mTmpDir).createSync(recursive: true);
  } catch (_) {}
}
final String cfgPath = '$baseDir/.app_cfg.json';
final String settingsPath = '$baseDir/A2R_Settings.xlsx';

String baseName(String path) => path.split('/').last;

List<Directory> findWhatsAppImageDirs() {
  final out = <Directory>[];

  void checkAndAdd(String p) {
    final d = Directory(p);
    if (d.existsSync() && !out.any((x) => x.path == d.path)) {
      out.add(d);
    }
  }

  checkAndAdd('$root/Android/media/com.whatsapp/WhatsApp/accounts/1003/Media/WhatsApp Images');
  checkAndAdd('$root/Android/media/com.whatsapp/WhatsApp/Media/WhatsApp Images');
  checkAndAdd('$root/WhatsApp/Media/WhatsApp Images');

  final mediaDir = Directory('$root/Android/media');
  if (mediaDir.existsSync()) {
    try {
      for (final pkg in mediaDir.listSync().whereType<Directory>()) {
        if (baseName(pkg.path).toLowerCase().contains('whatsapp')) {
          checkAndAdd('${pkg.path}/Media/WhatsApp Images');
          checkAndAdd('${pkg.path}/WhatsApp/Media/WhatsApp Images');
        }
      }
    } catch (_) {}
  }
  return out;
}

DateTime fileDate(File f) {
  final m = RegExp(r'-(\d{8})-').firstMatch(baseName(f.path));
  if (m != null) {
    final s = m.group(1)!;
    final d = DateTime.tryParse(
        '${s.substring(0, 4)}-${s.substring(4, 6)}-${s.substring(6, 8)}');
    if (d != null) return d;
  }
  return f.lastModifiedSync();
}

Map<String, dynamic> _readCfg() {
  try {
    return Map<String, dynamic>.from(jsonDecode(File(cfgPath).readAsStringSync()) as Map);
  } catch (_) {
    return <String, dynamic>{};
  }
}

void _writeCfg(Map<String, dynamic> m) {
  try {
    File(cfgPath)
      ..createSync(recursive: true)
      ..writeAsStringSync(jsonEncode(m));
  } catch (_) {}
}

String? loadFolder() => _readCfg()['folder'] as String?;
void saveFolder(String? f) {
  final m = _readCfg();
  m['folder'] = f;
  _writeCfg(m);
}
String? loadOutDir() => _readCfg()['out'] as String?;
void saveOutDir(String? f) {
  final m = _readCfg();
  m['out'] = f;
  _writeCfg(m);
}

// ───────────────── معالجة وتعديل الصور ─────────────────
bool _fixImage(String src, String dst, int deg) {
  final bytes = File(src).readAsBytesSync();
  var im = img.decodeImage(bytes);
  if (im == null) {
    if (src != dst) File(src).copySync(dst);
    return false;
  }
  var ori = 1;
  try {
    ori = im.exif.imageIfd.orientation ?? 1;
  } catch (_) {}
  if (deg == 0 && ori == 1) {
    if (src != dst) File(src).copySync(dst);
    return false;
  }
  im = img.bakeOrientation(im);
  if (deg != 0) im = img.copyRotate(im, angle: deg);
  File(dst).writeAsBytesSync(img.encodeJpg(im, quality: 92));
  return true;
}

void _enhanceImage(String src, String dst) {
  var im = img.decodeImage(File(src).readAsBytesSync());
  if (im == null) throw Exception('decode failed');
  im = img.bakeOrientation(im);
  im = img.grayscale(im);
  im = img.adjustColor(im, contrast: 1.6);
  if (im.width < 1800) {
    im = img.copyResize(im,
        width: (im.width * 1.6).round(), interpolation: img.Interpolation.cubic);
  } else if (im.width > 2600) {
    im = img.copyResize(im, width: 2200, interpolation: img.Interpolation.linear);
  }
  File(dst).writeAsBytesSync(img.encodeJpg(im, quality: 95));
}

Future<bool> fixImageAsync(String src, String dst, int deg) =>
    Isolate.run(() => _fixImage(src, dst, deg));

Future<void> enhanceImageAsync(String src, String dst) =>
    Isolate.run(() => _enhanceImage(src, dst));

// خوارزميات مسح الباتش الذكي
int _otsu(List<int> hist, int total) {
  var sumAll = 0;
  for (var i = 0; i < 256; i++) {
    sumAll += i * hist[i];
  }
  var wb = 0, sb = 0, th = 128;
  var best = -1.0;
  for (var t = 0; t < 256; t++) {
    wb += hist[t];
    if (wb == 0) continue;
    final wf = total - wb;
    if (wf == 0) break;
    sb += t * hist[t];
    final mb = sb / wb;
    final mf = (sumAll - sb) / wf;
    final v = wb.toDouble() * wf * (mb - mf) * (mb - mf);
    if (v > best) {
      best = v;
      th = t;
    }
  }
  return th;
}

List<double> _trimRange(List<double> v, double f) {
  final s = List<double>.from(v)..sort();
  final k = (s.length * f).floor();
  return [s[k], s[s.length - 1 - k]];
}

int _ci(int v, int lo, int hi) => v < lo ? lo : (v > hi ? hi : v);
double _cd(double v, double lo, double hi) => v < lo ? lo : (v > hi ? hi : v);

List<double>? _fitLine(List<double?> qs, List<double> ps, int skip) {
  var pp = <double>[];
  var qq = <double>[];
  for (var b = skip; b < qs.length - skip; b++) {
    final q = qs[b];
    if (q != null) {
      pp.add(ps[b]);
      qq.add(q);
    }
  }
  if (pp.length < 4) return null;
  var m = 0.0, k = 0.0;
  for (var iter = 0; iter < 2; iter++) {
    final n = pp.length;
    var sp = 0.0, sq = 0.0, spp = 0.0, spq = 0.0;
    for (var i = 0; i < n; i++) {
      sp += pp[i];
      sq += qq[i];
      spp += pp[i] * pp[i];
      spq += pp[i] * qq[i];
    }
    final d = n * spp - sp * sp;
    if (d.abs() < 1e-9) {
      m = 0.0;
      k = sq / n;
      break;
    }
    m = (n * spq - sp * sq) / d;
    k = (sq - m * sp) / n;
    final res = <double>[for (var i = 0; i < n; i++) (qq[i] - (m * pp[i] + k)).abs()]..sort();
    final lim = math.max(1.5, 2.5 * res[res.length ~/ 2]);
    final np = <double>[];
    final nq = <double>[];
    for (var i = 0; i < n; i++) {
      if ((qq[i] - (m * pp[i] + k)).abs() <= lim) {
        np.add(pp[i]);
        nq.add(qq[i]);
      }
    }
    if (np.length < 4) break;
    pp = np;
    qq = nq;
  }
  return [m, k];
}

List<double>? _quadOfComponent(List<int> pts, Int32List lab, int cid, int sw, int sh) {
  final bx = <double>[];
  final by = <double>[];
  for (final idx in pts) {
    final x = idx % sw, y = idx ~/ sw;
    var edge = false;
    if (x == 0 || x == sw - 1 || y == 0 || y == sh - 1) {
      edge = true;
    } else if (lab[idx - 1] != cid ||
        lab[idx + 1] != cid ||
        lab[idx - sw] != cid ||
        lab[idx + sw] != cid) {
      edge = true;
    }
    if (edge) {
      bx.add(x.toDouble());
      by.add(y.toDouble());
    }
  }
  if (bx.length < 20) return null;

  var bestArea = double.infinity;
  var bestDeg = 0;
  for (var deg = 0; deg < 90; deg += 2) {
    final t = deg * math.pi / 180;
    final c = math.cos(t), s = math.sin(t);
    final us = <double>[];
    final vs = <double>[];
    for (var i = 0; i < bx.length; i++) {
      us.add(bx[i] * c + by[i] * s);
      vs.add(-bx[i] * s + by[i] * c);
    }
    final ur = _trimRange(us, 0.006);
    final vr = _trimRange(vs, 0.006);
    final area = (ur[1] - ur[0]) * (vr[1] - vr[0]);
    if (area < bestArea) {
      bestArea = area;
      bestDeg = deg;
    }
  }
  final t0 = bestDeg * math.pi / 180;
  final eu = <double>[math.cos(t0), math.sin(t0)];
  final ev = <double>[-math.sin(t0), math.cos(t0)];
  final cands = <List<double>>[
    eu,
    <double>[-eu[0], -eu[1]],
    ev,
    <double>[-ev[0], -ev[1]],
  ];
  var ex = cands[0], ey = cands[0];
  for (final c in cands) {
    if (c[0] > ex[0]) ex = c;
    if (c[1] > ey[1]) ey = c;
  }
  final ax = <double>[];
  final ay = <double>[];
  for (var i = 0; i < bx.length; i++) {
    ax.add(bx[i] * ex[0] + by[i] * ex[1]);
    ay.add(bx[i] * ey[0] + by[i] * ey[1]);
  }
  final xr = _trimRange(ax, 0.006);
  final yr = _trimRange(ay, 0.006);
  final vertical = (yr[1] - yr[0]) >= (xr[1] - xr[0]);
  final lp = vertical ? ay : ax;
  final lq = vertical ? ax : ay;
  final l0 = vertical ? yr[0] : xr[0];
  final l1 = vertical ? yr[1] : xr[1];
  if (l1 - l0 < 4) return null;
  const nbins = 24;
  final lo = List<double?>.filled(nbins, null);
  final loP = List<double>.filled(nbins, 0.0);
  final hi = List<double?>.filled(nbins, null);
  final hiP = List<double>.filled(nbins, 0.0);
  for (var i = 0; i < lp.length; i++) {
    final p = lp[i], q = lq[i];
    if (p < l0 || p > l1) continue;
    final b = math.min(nbins - 1, ((p - l0) / (l1 - l0) * nbins).floor());
    final cl = lo[b];
    if (cl == null || q < cl) {
      lo[b] = q;
      loP[b] = p;
    }
    final ch = hi[b];
    if (ch == null || q > ch) {
      hi[b] = q;
      hiP[b] = p;
    }
  }
  final lowFit = _fitLine(lo, loP, 2) ?? <double>[0.0, vertical ? xr[0] : yr[0]];
  final highFit = _fitLine(hi, hiP, 2) ?? <double>[0.0, vertical ? xr[1] : yr[1]];
  double qLow(double p) => lowFit[0] * p + lowFit[1];
  double qHigh(double p) => highFit[0] * p + highFit[1];
  List<double> toImg(double a, double b) =>
      <double>[a * ex[0] + b * ey[0], a * ex[1] + b * ey[1]];
  final List<double> tl, tr, bl, br;
  if (vertical) {
    tl = toImg(qLow(l0), l0);
    tr = toImg(qHigh(l0), l0);
    bl = toImg(qLow(l1), l1);
    br = toImg(qHigh(l1), l1);
  } else {
    tl = toImg(l0, qLow(l0));
    bl = toImg(l0, qHigh(l0));
    tr = toImg(l1, qLow(l1));
    br = toImg(l1, qHigh(l1));
  }
  return <double>[tl[0], tl[1], tr[0], tr[1], bl[0], bl[1], br[0], br[1]];
}

double _quadArea(List<double> q) {
  final xs = [q[0], q[2], q[6], q[4]];
  final ys = [q[1], q[3], q[7], q[5]];
  var s = 0.0;
  for (var i = 0; i < 4; i++) {
    final j = (i + 1) % 4;
    s += xs[i] * ys[j] - xs[j] * ys[i];
  }
  return s.abs() / 2;
}

double _dist(double x1, double y1, double x2, double y2) =>
    math.sqrt((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2));

List<List<double>> _detectReceipts(img.Image im) {
  final w = im.width, h = im.height;
  final s = 480.0 / math.max(w, h);
  final sw = math.max(16, (w * s).round());
  final sh = math.max(16, (h * s).round());
  final small = img.copyResize(im, width: sw, height: sh, interpolation: img.Interpolation.linear);
  final total = sw * sh;
  final lum = Uint8List(total);
  final hist = List<int>.filled(256, 0);
  for (var y = 0; y < sh; y++) {
    for (var x = 0; x < sw; x++) {
      final p = small.getPixel(x, y);
      final l = _ci(((p.r * 299 + p.g * 587 + p.b * 114) ~/ 1000).toInt(), 0, 255);
      lum[y * sw + x] = l;
      hist[l]++;
    }
  }
  final th = _otsu(hist, total);
  final lab = Int32List(total);
  final comps = <List<int>>[];
  final stack = <int>[];
  var cid = 0;
  for (var i0 = 0; i0 < total; i0++) {
    if (lum[i0] <= th || lab[i0] != 0) continue;
    cid++;
    final pts = <int>[];
    lab[i0] = cid;
    stack.add(i0);
    while (stack.isNotEmpty) {
      final i = stack.removeLast();
      pts.add(i);
      final x = i % sw, y = i ~/ sw;
      if (x > 0 && lum[i - 1] > th && lab[i - 1] == 0) {
        lab[i - 1] = cid;
        stack.add(i - 1);
      }
      if (x < sw - 1 && lum[i + 1] > th && lab[i + 1] == 0) {
        lab[i + 1] = cid;
        stack.add(i + 1);
      }
      if (y > 0 && lum[i - sw] > th && lab[i - sw] == 0) {
        lab[i - sw] = cid;
        stack.add(i - sw);
      }
      if (y < sh - 1 && lum[i + sw] > th && lab[i + sw] == 0) {
        lab[i + sw] = cid;
        stack.add(i + sw);
      }
    }
    comps.add(pts);
  }
  if (comps.isEmpty) return [];
  final order = List<int>.generate(comps.length, (i) => i)
    ..sort((a, b) => comps[b].length.compareTo(comps[a].length));
  final mainSize = comps[order.first].length;
  if (mainSize < total * 0.04) return [];
  if (mainSize > total * 0.9) {
    return [
      [0.0, 0.0, w.toDouble(), 0.0, 0.0, h.toDouble(), w.toDouble(), h.toDouble()]
    ];
  }
  final out = <List<double>>[];
  for (final oi in order) {
    final pts = comps[oi];
    if (out.length >= 4) break;
    if (pts.length < mainSize * 0.3 || pts.length < total * 0.02) break;
    final q = _quadOfComponent(pts, lab, oi + 1, sw, sh);
    if (q == null) continue;
    final area = _quadArea(q);
    if (area <= 0 || pts.length / area < 0.45) continue;
    out.add(<double>[for (final v in q) v / s]);
  }
  out.sort((a, b) {
    final acx = (a[0] + a[2] + a[4] + a[6]) / 4, bcx = (b[0] + b[2] + b[4] + b[6]) / 4;
    final acy = (a[1] + a[3] + a[5] + a[7]) / 4, bcy = (b[1] + b[3] + b[5] + b[7]) / 4;
    final aw = _dist(a[0], a[1], a[2], a[3]), bw = _dist(b[0], b[1], b[2], b[3]);
    if ((acx - bcx).abs() > math.min(aw, bw) * 0.5) return acx.compareTo(bcx);
    return acy.compareTo(bcy);
  });
  return out;
}

img.Image? _warpAndEnhance(img.Image im, List<double> q) {
  final wq = (_dist(q[0], q[1], q[2], q[3]) + _dist(q[4], q[5], q[6], q[7])) / 2;
  final hq = (_dist(q[0], q[1], q[4], q[5]) + _dist(q[2], q[3], q[6], q[7])) / 2;
  if (wq < 40 || hq < 40) return null;
  final k = math.min(1.0, 2200.0 / math.max(wq, hq));
  final wo = math.max(8, (wq * k).round());
  final ho = math.max(8, (hq * k).round());

  final xs = [q[0], q[2], q[4], q[6]];
  final ys = [q[1], q[3], q[5], q[7]];
  final cx0 = _ci(xs.reduce(math.min).floor() - 6, 0, im.width - 2);
  final cy0 = _ci(ys.reduce(math.min).floor() - 6, 0, im.height - 2);
  final cx1 = _ci(xs.reduce(math.max).ceil() + 6, cx0 + 2, im.width);
  final cy1 = _ci(ys.reduce(math.max).ceil() + 6, cy0 + 2, im.height);
  final cw = cx1 - cx0, ch = cy1 - cy0;
  var crop = img.copyCrop(im, x: cx0, y: cy0, width: cw, height: ch);
  if (k < 1.0) {
    crop = img.copyResize(crop,
        width: math.max(4, (cw * k).round()),
        height: math.max(4, (ch * k).round()),
        interpolation: img.Interpolation.linear);
  }
  final gw = crop.width, gh = crop.height;
  final sx = gw / cw, sy = gh / ch;
  final g = Uint8List(gw * gh);
  for (var y = 0; y < gh; y++) {
    for (var x = 0; x < gw; x++) {
      final p = crop.getPixel(x, y);
      g[y * gw + x] = _ci(((p.r * 299 + p.g * 587 + p.b * 114) ~/ 1000).toInt(), 0, 255);
    }
  }
  final tlx = (q[0] - cx0) * sx, tly = (q[1] - cy0) * sy;
  final trx = (q[2] - cx0) * sx, try0 = (q[3] - cy0) * sy;
  final blx = (q[4] - cx0) * sx, bly = (q[5] - cy0) * sy;
  final brx = (q[6] - cx0) * sx, bry = (q[7] - cy0) * sy;

  const inset = 0.008;
  final og = Uint8List(wo * ho);
  for (var j = 0; j < ho; j++) {
    final v = inset + (1 - 2 * inset) * (j / (ho - 1));
    for (var i = 0; i < wo; i++) {
      final u = inset + (1 - 2 * inset) * (i / (wo - 1));
      final x = (1 - v) * ((1 - u) * tlx + u * trx) + v * ((1 - u) * blx + u * brx);
      final y = (1 - v) * ((1 - u) * tly + u * try0) + v * ((1 - u) * bly + u * bry);
      final x0 = _ci(x.floor(), 0, gw - 2);
      final y0 = _ci(y.floor(), 0, gh - 2);
      final fx = _cd(x - x0, 0.0, 1.0);
      final fy = _cd(y - y0, 0.0, 1.0);
      final a = g[y0 * gw + x0], b = g[y0 * gw + x0 + 1];
      final c = g[(y0 + 1) * gw + x0], d = g[(y0 + 1) * gw + x0 + 1];
      og[j * wo + i] =
          ((a * (1 - fx) + b * fx) * (1 - fy) + (c * (1 - fx) + d * fx) * fy).round();
    }
  }

  const bs = 12;
  final hb = (ho + bs - 1) ~/ bs, wb = (wo + bs - 1) ~/ bs;
  var bg = List<double>.filled(hb * wb, 0.0);
  for (var by = 0; by < hb; by++) {
    for (var bx = 0; bx < wb; bx++) {
      var mx = 0;
      final y1 = math.min(ho, (by + 1) * bs), x1 = math.min(wo, (bx + 1) * bs);
      for (var y = by * bs; y < y1; y++) {
        for (var x = bx * bs; x < x1; x++) {
          final vv = og[y * wo + x];
          if (vv > mx) mx = vv;
        }
      }
      bg[by * wb + bx] = mx.toDouble();
    }
  }
  for (var pass = 0; pass < 2; pass++) {
    final nbg = List<double>.filled(hb * wb, 0.0);
    for (var by = 0; by < hb; by++) {
      for (var bx = 0; bx < wb; bx++) {
        var sum = 0.0;
        for (var dy = -1; dy <= 1; dy++) {
          for (var dx = -1; dx <= 1; dx++) {
            sum += bg[_ci(by + dy, 0, hb - 1) * wb + _ci(bx + dx, 0, wb - 1)];
          }
        }
        nbg[by * wb + bx] = sum / 9;
      }
    }
    bg = nbg;
  }
  final res = img.Image(width: wo, height: ho);
  for (var y = 0; y < ho; y++) {
    final gy = _cd((y + 0.5) / bs - 0.5, 0.0, (hb - 1).toDouble());
    final by0 = gy.floor();
    final by1 = math.min(hb - 1, by0 + 1);
    final fy = gy - by0;
    for (var x = 0; x < wo; x++) {
      final gx = _cd((x + 0.5) / bs - 0.5, 0.0, (wb - 1).toDouble());
      final bx0 = gx.floor();
      final bx1 = math.min(wb - 1, bx0 + 1);
      final fx = gx - bx0;
      final top = bg[by0 * wb + bx0] * (1 - fx) + bg[by0 * wb + bx1] * fx;
      final bot = bg[by1 * wb + bx0] * (1 - fx) + bg[by1 * wb + bx1] * fx;
      final bgv = top * (1 - fy) + bot * fy;
      final n = _cd(og[y * wo + x] / math.max(bgv, 40.0) * 255.0, 0.0, 255.0);
      final o = math.pow(_cd((n - 30.0) / (240.0 - 30.0), 0.0, 1.0), 1.6).toDouble();
      final c = (o * 255).round();
      res.setPixelRgb(x, y, c, c, c);
    }
  }
  return res;
}

List<String> _scanToPages(String src, String base) {
  final cached = <String>[];
  var ci = 0;
  while (File('${base}_$ci.jpg').existsSync()) {
    cached.add('${base}_$ci.jpg');
    ci++;
  }
  if (cached.isNotEmpty) return cached;
  if (File('${base}_none').existsSync()) return <String>[];
  final out = <String>[];
  try {
    var im = img.decodeImage(File(src).readAsBytesSync());
    if (im != null) {
      im = img.bakeOrientation(im);
      final quads = _detectReceipts(im);
      for (final q in quads) {
        final o = _warpAndEnhance(im, q);
        if (o == null) continue;
        final p = '${base}_${out.length}.jpg';
        File(p).writeAsBytesSync(img.encodeJpg(o, quality: 90));
        out.add(p);
      }
    }
  } catch (_) {}
  if (out.isEmpty) {
    try {
      File('${base}_none').writeAsStringSync('1');
    } catch (_) {}
  }
  return out;
}

Future<List<String>> scanReceiptAsync(String src) async {
  try {
    final st = File(src).statSync();
    Directory('$baseDir/.scans').createSync(recursive: true);
    final safe = baseName(src).replaceAll(RegExp(r'[^A-Za-z0-9_\-]'), '_');
    final base = '$baseDir/.scans/v2_${safe}_${st.size}_${st.modified.millisecondsSinceEpoch}';
    return await Isolate.run(() => _scanToPages(src, base));
  } catch (_) {
    return <String>[];
  }
}

// ───────────────── ملء شيت الإكسيل ─────────────────
Future<List<int>?> loadTemplateBytes() async {
  try {
    final mine = File('$baseDir/template.xlsx');
    if (mine.existsSync()) return mine.readAsBytesSync();
  } catch (_) {}
  try {
    final data = await rootBundle.load('assets/template.xlsx');
    return data.buffer.asUint8List();
  } catch (_) {}
  return _blankTemplateBytes();
}

String _firstSheetPath(Archive arc) {
  String? text(String n) {
    for (final f in arc.files) {
      if (f.name == n) return utf8.decode(f.content as List<int>, allowMalformed: true);
    }
    return null;
  }

  final wb = text('xl/workbook.xml');
  final rels = text('xl/_rels/workbook.xml.rels');
  if (wb != null && rels != null) {
    final sm = RegExp(r'<sheet\b[^>]*?r:id="([^"]+)"').firstMatch(wb);
    if (sm != null) {
      for (final m in RegExp(r'<Relationship\b[^>]*>').allMatches(rels)) {
        final tag = m.group(0)!;
        if (tag.contains('Id="${sm.group(1)}"')) {
          final t = RegExp(r'Target="([^"]+)"').firstMatch(tag)?.group(1);
          if (t != null) return t.startsWith('/') ? t.substring(1) : 'xl/$t';
        }
      }
    }
  }
  return 'xl/worksheets/sheet1.xml';
}

final RegExp _cellRe = RegExp(r'<c\b[^>]*?/>|<c\b[^>]*?>.*?</c>', dotAll: true);

int _colNum(String c) {
  var n = 0;
  for (final ch in c.codeUnits) {
    n = n * 26 + (ch - 64);
  }
  return n;
}

RegExpMatch? _findRow(String xml, int row) => RegExp(
        '<row\\b[^>]*\\br="$row"[^>]*?(?:/>|>.*?</row>)',
        dotAll: true)
    .firstMatch(xml);

class _UnitOut {
  final String name;
  final int count;
  final double amount;
  _UnitOut(this.name, this.count, this.amount);
}

String _xmlEsc(String s) =>
    s.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');

String _numStr(double v) {
  final r = double.parse(v.toStringAsFixed(2));
  return r == r.roundToDouble() ? r.toInt().toString() : r.toString();
}

String buildUnitsSheetXml(String xml, List<_UnitOut> units) {
  final sd = RegExp(r'<sheetData\b[^>]*?(?:/>|>(.*?)</sheetData>)', dotAll: true).firstMatch(xml);
  if (sd == null) throw Exception('sheetData missing');
  final inner = sd.group(1) ?? '';
  final rowRe = RegExp(r'<row\b[^>]*?(?:/>|>.*?</row>)', dotAll: true);
  final header = StringBuffer();
  String? proto;
  for (final m in rowRe.allMatches(inner)) {
    final row = m.group(0)!;
    final n = int.tryParse(RegExp(r'\br="(\d+)"').firstMatch(row)?.group(1) ?? '') ?? 0;
    if (n <= 1) {
      header.write(row);
    } else {
      proto ??= row;
    }
  }
  final cellRe = RegExp(r'<c\b[^>]*?/>|<c\b[^>]*?>.*?</c>', dotAll: true);
  final protoCells = <String, String>{};
  var rowOpen = '';
  if (proto != null) {
    rowOpen = RegExp(r'<row\b[^>]*?(?=/?>)').firstMatch(proto)!.group(0)!;
    for (final m in cellRe.allMatches(proto)) {
      final cell = m.group(0)!;
      final open = RegExp(r'<c\b[^>]*?/?>').firstMatch(cell)!.group(0)!;
      final col = RegExp(r'\br="([A-Z]+)\d+"').firstMatch(open)?.group(1);
      if (col != null) protoCells[col] = cell;
    }
  }
  String styleOf(String col) {
    final c = protoCells[col];
    if (c == null) return '';
    final open = RegExp(r'<c\b[^>]*?/?>').firstMatch(c)!.group(0)!;
    final st = RegExp(r'\bs="(\d+)"').firstMatch(open)?.group(1);
    return st == null ? '' : ' s="$st"';
  }

  final cols = <String>{...protoCells.keys, 'B', 'C', 'D', 'F'}.toList()
    ..sort((a, b) => _colNum(a).compareTo(_colNum(b)));
  final body = StringBuffer();
  for (var i = 0; i < units.length; i++) {
    final r = i + 2;
    final u = units[i];
    final open = proto == null
        ? '<row r="$r"'
        : rowOpen.replaceFirst(RegExp(r'\br="\d+"'), 'r="$r"');
    body.write('$open>');
    for (final col in cols) {
      final st = styleOf(col);
      switch (col) {
        case 'B':
          body.write('<c r="B$r"$st t="inlineStr"><is><t xml:space="preserve">${_xmlEsc(u.name)}</t></is></c>');
          break;
        case 'C':
          body.write('<c r="C$r"$st><v>${u.count}</v></c>');
          break;
        case 'D':
          body.write('<c r="D$r"$st><v>${_numStr(u.amount)}</v></c>');
          break;
        case 'E':
          body.write('<c r="E$r"$st/>');
          break;
        case 'F':
          body.write('<c r="F$r"$st><f>D$r+IF(ISBLANK(E$r),0,E$r)</f><v>${_numStr(u.amount)}</v></c>');
          break;
        default:
          final c = protoCells[col];
          if (c == null || c.contains('<f')) {
            body.write('<c r="$col$r"$st/>');
          } else {
            body.write(c.replaceFirst(RegExp(r'\br="[A-Z]+\d+"'), 'r="$col$r"'));
          }
      }
    }
    body.write('</row>');
  }
  var out = xml.replaceRange(sd.start, sd.end, '<sheetData>${header.toString()}${body.toString()}</sheetData>');
  final lastCol = cols.last;
  out = out.replaceFirst(RegExp(r'<dimension\b[^>]*?/>'), '<dimension ref="A1:$lastCol${units.length + 1}"/>');
  return out;
}

List<int> fillTemplate(List<int> bytes, List<_UnitOut> units) {
  final arc = ZipDecoder().decodeBytes(bytes);
  final sheetPath = _firstSheetPath(arc);
  final out = Archive();
  var found = false;
  for (final f in arc.files) {
    if (!f.isFile) continue;
    if (f.name == 'xl/calcChain.xml') continue;
    final c = f.content as List<int>;
    if (f.name == sheetPath) {
      found = true;
      final xml = buildUnitsSheetXml(utf8.decode(c, allowMalformed: true), units);
      final b = utf8.encode(xml);
      out.addFile(ArchiveFile(f.name, b.length, b));
    } else if (f.name == '[Content_Types].xml' || f.name == 'xl/_rels/workbook.xml.rels') {
      final t = utf8
          .decode(c, allowMalformed: true)
          .replaceAll(RegExp(r'<Override\b[^>]*calcChain[^>]*?/>'), '')
          .replaceAll(RegExp(r'<Relationship\b[^>]*calcChain[^>]*?/>'), '');
      final b = utf8.encode(t);
      out.addFile(ArchiveFile(f.name, b.length, b));
    } else {
      out.addFile(ArchiveFile(f.name, c.length, c));
    }
  }
  if (!found) throw Exception('sheet not found: $sheetPath');
  final enc = ZipEncoder().encode(out);
  if (enc == null) throw Exception('zip encode failed');
  return enc;
}

const MethodChannel _chan = MethodChannel('pos/native');

String todayToken() {
  final n = DateTime.now();
  return '${n.day}.${n.month}';
}

String _normAr(String s) => s
    .replaceAll(RegExp('[أإآ]'), 'ا')
    .replaceAll('ة', 'ه')
    .replaceAll('ى', 'ي')
    .toLowerCase();

// ───────────────────── نموذج الإيصال المحسّن ─────────────────────
class ReceiptItem {
  String path;
  String? tid;
  int? orderNum;
  int? row;
  String? unitName;
  double amount;
  String? date;
  String key;
  bool ok;
  String how;
  List<double> batches;
  int? suffix;
  String? origTid;
  bool conflict;
  bool skip;
  double? alt;
  bool dupSuspect;
  bool userExcluded;
  String? dupDecision;
  int? dupOf;
  String? unusual;
  bool ignoredDoc;
  String? ign;
  bool dateOdd;
  String? dateMain;
  String? dateFix;
  int? txn;
  bool reprintSummary;
  bool missingBatchClosed;
  List<double> cands = <double>[];

  ReceiptItem({
    required this.path,
    this.tid,
    this.orderNum,
    this.row,
    this.unitName,
    this.amount = 0.0,
    this.date,
    this.key = '',
    this.ok = false,
    this.how = 'none',
    List<double>? batches,
    this.suffix,
    this.origTid,
    this.conflict = false,
    this.skip = false,
    this.alt,
    this.dupSuspect = false,
    this.userExcluded = false,
    this.dupDecision,
    this.dupOf,
    this.ignoredDoc = false,
    this.ign,
    this.dateOdd = false,
    this.reprintSummary = false,
    this.missingBatchClosed = false,
  }) : batches = batches ?? <double>[];

  bool get counts => !skip && !dupSuspect && !userExcluded && !ignoredDoc;
  bool get weakId =>
      !(how == 'terminal' || how == 'fuzzy' || how == 'manual' || how == 'extra-batch');

  bool get needsReview =>
      ignoredDoc ||
      dupSuspect ||
      dateOdd ||
      reprintSummary ||
      missingBatchClosed ||
      (!userExcluded &&
          (amount <= 0 ||
              row == null ||
              !ok ||
              alt != null ||
              unusual != null ||
              weakId));
}

final String fixPath = '$baseDir/.app_fix.json';

Map<String, dynamic> loadFixes() {
  try {
    return Map<String, dynamic>.from(jsonDecode(File(fixPath).readAsStringSync()) as Map);
  } catch (_) {
    return <String, dynamic>{};
  }
}

void saveFix(ReceiptItem item) {
  if (item.key.isEmpty) return;
  try {
    final m = loadFixes();
    m[item.key] = {
      'val': item.amount,
      'tid': item.tid,
      if (item.dupDecision != null) 'dup': item.dupDecision,
      if (item.ign != null) 'ign': item.ign,
      if (item.dateFix != null) 'date': item.dateFix,
      'rep': item.reprintSummary,
      'nobc': item.missingBatchClosed,
    };
    File(fixPath)
      ..createSync(recursive: true)
      ..writeAsStringSync(jsonEncode(m));
  } catch (_) {}
}

void keepIgnored(ReceiptItem it) {
  try {
    if (!it.path.startsWith('$rawImagesDir/')) {
      Directory(rawImagesDir).createSync(recursive: true);
      final dst = '$rawImagesDir/wa_${baseName(it.path)}';
      if (!File(dst).existsSync()) File(it.path).copySync(dst);
      it.path = dst;
    }
  } catch (_) {}
  it.ignoredDoc = false;
  it.userExcluded = false;
  it.ign = 'keep';
  saveFix(it);
}

void keepAsExtraBatch(ReceiptItem it, List<ReceiptItem> all) {
  var k = 2;
  bool used(int n) => all.any((x) =>
      !identical(x, it) && !x.skip && x.counts && x.orderNum == it.orderNum && (x.suffix ?? 1) == n);
  while (used(k)) {
    k++;
  }
  it.suffix = k;
  it.how = 'extra-batch';
  it.dupSuspect = false;
  it.userExcluded = false;
  it.dupDecision = 'keep';
  try {
    final cut = it.path.lastIndexOf('/');
    final dot = it.path.lastIndexOf('.');
    final ext = dot > cut ? it.path.substring(dot) : '.jpg';
    final dst = '${it.path.substring(0, cut + 1)}${it.orderNum}_$k$ext';
    if (dst != it.path && !File(dst).existsSync()) {
      it.path = File(it.path).renameSync(dst).path;
    }
  } catch (_) {}
  saveFix(it);
}

void checkUnusual(List<ReceiptItem> items, String reportDate) {
  for (final r in items) {
    r.unusual = null;
  }
  try {
    final cfg = _readCfg();
    final hist = <String, Map<String, dynamic>>{};
    final raw = cfg['hist'];
    if (raw is Map) {
      raw.forEach((k, v) {
        if (v is Map) hist[k.toString()] = Map<String, dynamic>.from(v);
      });
    }
    final today = <int, double>{};
    for (final r in items) {
      if (r.orderNum != null && r.amount > 0) {
        today[r.orderNum!] = (today[r.orderNum!] ?? 0) + r.amount;
      }
    }
    for (final e in today.entries) {
      final prev = <double>[];
      hist.forEach((d, m) {
        if (d == reportDate) return;
        final v = m['${e.key}'];
        if (v is num && v > 0) prev.add(v.toDouble());
      });
      if (prev.length < 3) continue;
      final avg = prev.reduce((a, b) => a + b) / prev.length;
      if (avg < 1) continue;
      if (e.value > avg * 2.5 || e.value < avg / 2.5) {
        final msg = 'إجمالي ماكينة ${e.key} (${e.value.toStringAsFixed(2)}) بعيد عن متوسط '
            'الأيام اللي فاتت (${avg.toStringAsFixed(0)}) - اتأكد من المبلغ';
        for (final r in items) {
          if (r.orderNum == e.key && r.how != 'manual') r.unusual = msg;
        }
      }
    }
    if (reportDate != 'اليوم') {
      hist[reportDate] = {for (final e in today.entries) '${e.key}': e.value};
      while (hist.length > 30) {
        hist.remove(hist.keys.first);
      }
      cfg['hist'] = hist;
      _writeCfg(cfg);
    }
  } catch (_) {}
}

// ───────── خوارزميات عد إيصالات الدفع المزدوجة والمتوازية ─────────
const int _kParWorkers = 2;
bool fastCount = true;

bool _loadDual() {
  try {
    return _readCfg()['dualCount'] != false;
  } catch (_) {
    return true;
  }
}

bool dualCount = _loadDual();

void toggleDualCount() {
  dualCount = !dualCount;
  try {
    final m = _readCfg();
    m['dualCount'] = dualCount;
    _writeCfg(m);
  } catch (_) {}
}

List<List<int>> _pdfJpegRanges(Uint8List b) {
  final out = <List<int>>[];
  int find(String pat, int from) {
    final p = pat.codeUnits;
    final n = b.length - p.length;
    for (var i = from; i <= n; i++) {
      if (b[i] != p[0]) continue;
      var ok = true;
      for (var j = 1; j < p.length; j++) {
        if (b[i + j] != p[j]) {
          ok = false;
          break;
        }
      }
      if (ok) return i;
    }
    return -1;
  }

  var pos = 0;
  while (true) {
    final i = find('/DCTDecode', pos);
    if (i < 0) break;
    final sIdx = find('stream', i);
    if (sIdx < 0) break;
    var st = sIdx + 6;
    if (st < b.length && b[st] == 13) st++;
    if (st < b.length && b[st] == 10) st++;
    final e = find('endstream', st);
    if (e < 0) break;
    var en = e;
    while (en > st && (b[en - 1] == 10 || b[en - 1] == 13 || b[en - 1] == 32)) {
      en--;
    }
    if (en - st > 20000 && b[st] == 0xFF && b[st + 1] == 0xD8 && b[en - 2] == 0xFF && b[en - 1] == 0xD9) {
      out.add([st, en]);
    }
    pos = e + 9;
  }
  return out;
}

class _SA {
  final String t;
  final double cx, cy, l, r;
  _SA(this.t, this.cx, this.cy, this.l, this.r);
}

class _StructPage {
  int sm = 0, mx = 0;
  int sm2 = 0;
  int sm3 = 0;
  double h = 3000;
  final Map<String, int> types = {};
  final List<String> dbgAmt = [];
}

int _seqCountCol(List<_SA> col, double h) {
  const rank = <String, int>{'logo': 0, 'term': 1, 'rrn': 2, 'amount': 3, 'appr': 4};
  final c = [...col]..sort((a, b) => a.cy.compareTo(b.cy));
  final strong = <double>[
    for (final a in c)
      if (a.t == 'logo' || a.t == 'amount' || a.t == 'appr') a.cy
  ]..sort();
  final gaps = <double>[];
  for (var i = 1; i < strong.length; i++) {
    final g = strong[i] - strong[i - 1];
    if (g > 0.03 * h) gaps.add(g);
  }
  var pitch = 0.1 * h;
  if (gaps.isNotEmpty) {
    gaps.sort();
    final m = gaps.length ~/ 2;
    pitch = gaps.length.isOdd ? gaps[m] : (gaps[m - 1] + gaps[m]) / 2;
  }
  final minP = 0.7 * pitch;
  var n = 0, last = -1;
  final lastCy = <String, double>{};
  for (final a in c) {
    final rk = rank[a.t] ?? 0;
    if (last < 0) {
      n = 1;
    } else if (rk <= last) {
      final prev = lastCy[a.t];
      if (prev != null && a.cy - prev < minP) continue;
      n++;
    }
    last = rk;
    lastCy[a.t] = a.cy;
  }
  return n;
}

bool _isIdSlip(RecognizedText res) {
  final lines = toLines(res);
  if (lines.isEmpty || lines.length > 16) return false;
  final all = lines.map((l) => l.text.toUpperCase()).join('\n');
  if (RegExp(r'TERMINAL|RRN|AMOUNT|APPROVED|NATIONAL|BANK\s*OF\s*EGYPT|BATCH|TOTAL|SALE|GRAND|CARD|REPRINT')
      .hasMatch(all)) {
    return false;
  }
  if (RegExp(r'(?<!\d)\d{12}(?!\d)').hasMatch(all.replaceAll(RegExp(r'[ \t]'), ''))) return false;
  final hasTid = RegExp(r'\bT\s?[I1L|]\s?D\b\W{0,4}[\dOo]{7,9}').hasMatch(all) ||
      RegExp(r'\bT\s?[I1L|]\s?D\b').hasMatch(all);
  if (hasTid) return true;
  final hasDate = RegExp(r'\d{2}\s?[/.\-]\s?\d{2}\s?[/.\-]\s?20\d{2}').hasMatch(all);
  return hasDate && lines.length <= 8 && findTerminalIn([all]) != null;
}

_StructPage _structCount(RecognizedText res) {
  final out = _StructPage();
  final lines = toLines(res);
  if (lines.isEmpty) return out;
  var h = 0.0, w = 0.0;
  for (final l in lines) {
    if (l.bottom > h) h = l.bottom;
    if (l.right > w) w = l.right;
  }
  out.h = h;
  final all = <_SA>[];
  final nameRe = RegExp(r'NATIONAL|BANK\s*OF\s*EGYPT');
  final rrnRe = RegExp(r'(?<![A-Z])R{1,2}N\W{0,3}\d');
  final amtRe = RegExp(r'^\s*AMOUNT\b');
  final amtBad = RegExp(r'FROM|ACCOUNT|CONFIRM|DEBIT|\bMY\b');
  for (final l in lines) {
    final u = l.text.toUpperCase();
    void add(String t) => all.add(_SA(t, (l.left + l.right) / 2, l.cy, l.left, l.right));
    if (nameRe.hasMatch(u)) add('logo');
    if (u.contains('TERMINAL')) add('term');
    if (rrnRe.hasMatch(u)) add('rrn');
    if (amtRe.hasMatch(u) && !amtBad.hasMatch(u)) {
      add('amount');
      out.dbgAmt.add('${l.text.length > 22 ? l.text.substring(0, 22) : l.text}@${h > 0 ? (l.cy * 100 / h).round() : 0}%');
    }
    if (u.contains('APPROVED')) add('appr');
  }

  final ded = <_SA>[];
  for (final a in all) {
    var dup = false;
    for (final o in ded) {
      if (o.t == a.t && o.l < a.r && a.l < o.r && (o.cy - a.cy).abs() < h * 0.02) {
        dup = true;
        break;
      }
    }
    if (!dup) ded.add(a);
  }
  for (final a in ded) {
    out.types[a.t] = (out.types[a.t] ?? 0) + 1;
  }
  if (ded.isEmpty) return out;
  out.mx = out.types.values.reduce((a, b) => a > b ? a : b);

  ded.sort((a, b) => a.cx.compareTo(b.cx));
  final cols = <List<_SA>>[];
  var cur = <_SA>[ded.first];
  for (var i = 1; i < ded.length; i++) {
    if (ded[i].cx - cur.last.cx > w * 0.07) {
      cols.add(cur);
      cur = <_SA>[];
    }
    cur.add(ded[i]);
  }
  cols.add(cur);

  var n = 0, n2 = 0, n3 = 0;
  for (final c in cols) {
    final cnt = <String, int>{};
    for (final a in c) {
      cnt[a.t] = (cnt[a.t] ?? 0) + 1;
    }
    n += cnt.values.reduce((a, b) => a > b ? a : b);
    var pay = 0;
    for (final k in const ['rrn', 'amount', 'appr']) {
      final v = cnt[k] ?? 0;
      if (v > pay) pay = v;
    }
    n2 += pay;
    n3 += _seqCountCol(c, h);
  }
  out.sm = n;
  out.sm2 = n2;
  out.sm3 = n3;
  return out;
}

bool sameDay(ReceiptItem a, ReceiptItem b) =>
    a.date == null || b.date == null || a.date == b.date;

Future<bool> chooseAndOpenFolder(BuildContext context, String path) async {
  try {
    final ok = await _chan.invokeMethod<bool>('openFolder', {'path': path});
    return ok == true;
  } catch (_) {
    return false;
  }
}

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

void resetDefaultHandlers(BuildContext context) {
  final cfg = _readCfg();
  cfg.remove('defFolder');
  cfg.remove('defImage');
  cfg.remove('defPdf');
  _writeCfg(cfg);
  ScaffoldMessenger.of(context).showSnackBar(
    const SnackBar(content: w.Text('تم مسح التطبيقات الافتراضية.')),
  );
}

Widget _statCard(IconData icon, String label, String value, Color c) => Expanded(
      child: Card(
        margin: const EdgeInsets.symmetric(horizontal: 2),
        child: Padding(
          padding: const EdgeInsets.symmetric(vertical: 8, horizontal: 4),
          child: Column(
            children: [
              Icon(icon, size: 18, color: c),
              const SizedBox(height: 2),
              FittedBox(
                  child: w.Text(value, style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 15))),
              Text(label, style: const TextStyle(fontSize: 11)),
            ],
          ),
        ),
      ),
    );

Color _unitColor(int? row) {
  const p = [
    Colors.blue, Colors.teal, Colors.purple, Colors.orange, Colors.pink, Colors.indigo,
    Colors.green, Colors.brown, Colors.cyan, Colors.red, Colors.lime, Colors.deepPurple,
  ];
  if (row == null) return Colors.grey;
  return p[row % p.length];
}

Future<void> openFileStatic(BuildContext context, String? path, String mime) async {
  if (path == null || !File(path).existsSync()) return;
  try {
    await _chan.invokeMethod<bool>('openFile', {'path': path, 'mime': mime});
  } catch (_) {}
}

// ───────────────────── شاشة التقفيل ─────────────────────
class HomePage extends StatefulWidget {
  const HomePage({super.key});

  @override
  State<HomePage> createState() => _HomePageState();
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

  Future<void> _refresh() async {
    if (busy) {
      if (mounted) setState(() {});
    } else {
      await rebuildFiles();
    }
  }

  Future<void> _onSwipe(ReceiptItem it, bool approve) async {
    if (approve) {
      if (it.dupSuspect) {
        keepAsExtraBatch(it, processedItems);
        await _refresh();
      } else if (it.dateOdd) {
        it.dateFix = it.date;
        it.dateOdd = false;
        saveFix(it);
        await _refresh();
      } else {
        if (it.amount <= 0 || it.row == null) {
          ScaffoldMessenger.of(context).showSnackBar(SnackBar(
              content: w.Text(_l(context, 'كمّل المبلغ والماكينة من زر القلم الأول.',
                  'Fill in the amount and machine with the pencil button first.'))));
          return;
        }
        it.ok = true;
        it.alt = null;
        it.unusual = null;
        it.userExcluded = false;
        it.ignoredDoc = false;
        it.reprintSummary = false;
        it.missingBatchClosed = false;
        if (it.how != 'extra-batch') it.how = 'manual';
        saveFix(it);
        await _refresh();
      }
    } else {
      if (it.dupSuspect) {
        it.dupSuspect = false;
        it.userExcluded = true;
        it.dupDecision = 'drop';
        saveFix(it);
        await _refresh();
      } else if (it.dateOdd) {
        if (it.dateMain != null) it.date = it.dateMain;
        it.dateFix = it.date;
        it.dateOdd = false;
        saveFix(it);
        await _refresh();
      } else {
        it.userExcluded = true;
        it.ok = false;
        saveFix(it);
        await _refresh();
      }
    }
  }

  Widget _swipeWrap(ReceiptItem item, Widget card) {
    final okLabel = item.dupSuspect
        ? _l(context, 'احتسبه باتش ✓', 'Count it ✓')
        : (item.dateOdd ? _l(context, 'سيب تاريخه ✓', 'Keep date ✓') : _l(context, 'صح ✓', 'OK ✓'));
    final badLabel = item.dupSuspect
        ? _l(context, 'مكرر - استبعده ✗', 'Duplicate ✗')
        : (item.dateOdd ? _l(context, 'تاريخ الأغلب ✗', 'Use main date ✗') : _l(context, 'غلط (استبعاد) ✗', 'Wrong (Exclude) ✗'));
    return Dismissible(
      key: ValueKey('sw_${item.path}_${item.userExcluded}_${item.ok}'),
      direction: DismissDirection.horizontal,
      confirmDismiss: (dir) async {
        await _onSwipe(item, dir == DismissDirection.startToEnd);
        return false;
      },
      background: Container(
        margin: const EdgeInsets.symmetric(vertical: 4),
        color: _tintBg(context, Colors.green, strong: true),
        alignment: AlignmentDirectional.centerStart,
        padding: const EdgeInsets.symmetric(horizontal: 20),
        child: w.Text(okLabel, style: const TextStyle(fontWeight: FontWeight.bold)),
      ),
      secondaryBackground: Container(
        margin: const EdgeInsets.symmetric(vertical: 4),
        color: _tintBg(context, Colors.red, strong: true),
        alignment: AlignmentDirectional.centerEnd,
        padding: const EdgeInsets.symmetric(horizontal: 20),
        child: w.Text(badLabel, style: const TextStyle(fontWeight: FontWeight.bold)),
      ),
      child: card,
    );
  }

  @override
  void didChangeAppLifecycleState(AppLifecycleState state) {
    _life = state;
    if (state == AppLifecycleState.resumed) prepareDefaultFiles();
  }

  @override
  void dispose() {
    _autoTimer?.cancel();
    homeRunG = null;
    homeSumG = null;
    homeAutoG = false;
    sharedDays.removeListener(_syncDays);
    wipeSignal.removeListener(_onWipe);
    WidgetsBinding.instance.removeObserver(this);
    super.dispose();
  }

  String _scanSig() {
    try {
      final dirs = customDir != null ? [Directory(customDir!)] : findWhatsAppImageDirs();
      final now = DateTime.now();
      final cutoff = days < 0
          ? DateTime(2000)
          : DateTime(now.year, now.month, now.day).subtract(Duration(days: days));
      var n = 0, size = 0, latest = 0;
      for (final d in dirs) {
        for (final e in d.listSync().whereType<File>()) {
          final lower = e.path.toLowerCase();
          if (!(lower.endsWith('.jpg') ||
              lower.endsWith('.jpeg') ||
              lower.endsWith('.png') ||
              lower.endsWith('.webp'))) {
            continue;
          }
          final b = baseName(e.path);
          if (b.startsWith('.') || b.startsWith('dup_')) continue;
          final fd = fileDate(e);
          if (fd.isBefore(cutoff)) continue;
          n++;
          size += e.lengthSync();
          final ms = fd.millisecondsSinceEpoch;
          if (ms > latest) latest = ms;
        }
      }
      return n == 0 ? '' : '$n:$size:$latest';
    } catch (_) {
      return '';
    }
  }

  Future<void> _watchService(bool on) async {
    try {
      if (on) {
        await _chan.invokeMethod('startKeepAlive',
            {'title': 'A2R POS', 'text': 'مراقبة الإيصالات الجديدة...'});
      } else {
        await _chan.invokeMethod('stopKeepAlive');
      }
    } catch (_) {}
  }

  void _startAuto() {
    _autoTimer?.cancel();
    _autoTimer = Timer.periodic(const Duration(seconds: 15), (_) => _autoTick());
    _watchService(true);
  }

  void _stopAuto() {
    _autoTimer?.cancel();
    _autoTimer = null;
    _watchService(false);
  }

  Future<void> _autoTick() async {
    if (!mounted || !autoScan || busy || chainRunning || machBusyG) return;
    final sig = _scanSig();
    if (sig.isEmpty || sig == _doneSig) {
      _pendSig = '';
      return;
    }
    if (sig != _pendSig) {
      _pendSig = sig;
      return;
    }
    _pendSig = '';
    await runFullPipeline(auto: true);
  }

  Future<void> _setAuto(bool v) async {
    homeAutoG = v;
    setState(() => autoScan = v);
    final cfg = _readCfg();
    cfg['autoScan'] = v;
    _writeCfg(cfg);
    if (v) {
      _doneSig = '';
      _startAuto();
      await askBatteryExempt();
    } else {
      _stopAuto();
    }
  }

  String _periodLabel() => days == 0 ? 'اليوم' : days == 1 ? 'يومين' : days == 2 ? '3 أيام' : 'الكل';

  Widget _statsRow() {
    final counted = processedItems
        .where((e) => !e.userExcluded && !e.ignoredDoc && e.ign != 'drop' && e.amount > 0)
        .toList();
    final total = counted.fold<double>(0, (a, e) => a + e.amount);
    final rv = processedItems.where((e) => e.needsReview).length;
    return Row(
      children: [
        _statCard(Icons.receipt_long, 'إيصالات', '${counted.length}', Colors.blue),
        _statCard(Icons.payments, 'الإجمالي', total.toStringAsFixed(2), Colors.green),
        _statCard(Icons.warning_amber, 'ناقص', '${missingTerminals.length}', Colors.amber),
        _statCard(Icons.fact_check, 'مراجعة', '$rv', Colors.orange),
      ],
    );
  }

  Future<void> pickFolder() async {
    final perm = await Permission.manageExternalStorage.request();
    if (!perm.isGranted) return;
    if (!mounted) return;
    final res = await Navigator.push<String>(
      context,
      MaterialPageRoute(builder: (_) => FolderPicker(start: customDir ?? root)),
    );
    if (res != null) {
      saveFolder(res);
      setState(() => customDir = res);
    }
  }

  Future<void> pickOutFolder() async {
    final perm = await Permission.manageExternalStorage.request();
    if (!perm.isGranted) return;
    if (!mounted) return;
    final res = await Navigator.push<String>(
      context,
      MaterialPageRoute(builder: (_) => FolderPicker(start: customOut ?? root)),
    );
    if (res != null) {
      saveOutDir(res);
      setState(() {
        customOut = res;
        outputDir = res;
      });
    }
  }

  Future<void> pickManualImages() async {
    try {
      final result = await FilePicker.pickFiles(type: FileType.image);
      if (result.isNotEmpty) {
        final paths = result.map((f) => f.path).whereType<String>().toList();
        if (paths.isNotEmpty) await runFullPipeline(manualPaths: paths);
      }
    } catch (_) {}
  }

  void useDefaultOut() {
    saveOutDir(null);
    setState(() {
      customOut = null;
      outputDir = defaultOutputDir;
    });
  }

  Future<void> openFolder(String path) async {
    try {
      Directory(path).createSync(recursive: true);
    } catch (_) {}
    try {
      if (await chooseAndOpenFolder(context, path)) return;
    } catch (_) {}
    await Clipboard.setData(ClipboardData(text: path));
    if (!mounted) return;
    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(content: Text('المسار: $path')),
    );
  }

  Future<void> sendToWhatsApp() async {
    final x = latestOutput('.xlsx');
    final p = latestOutput('.pdf');
    final files = <XFile>[if (x != null) XFile(x), if (p != null) XFile(p)];
    if (files.isEmpty) return;
    if (!await confirmSend(context,
        review: processedItems.where((e) => e.needsReview).length, missing: missingTerminals.length)) {
      return;
    }
    await Share.shareXFiles(files, text: 'تقرير الإيصالات');
  }

  String? latestOutput(String ext) {
    final cur = ext == '.xlsx' ? standardXlsxPath : generatedPdfPath;
    if (cur != null && File(cur).existsSync()) return cur;
    try {
      final d = Directory(outputDir);
      if (!d.existsSync()) return null;
      final fs = d
          .listSync()
          .whereType<File>()
          .where((f) => _isOutputFileName(baseName(f.path)) && f.path.endsWith(ext))
          .toList()
        ..sort((a, b) => b.lastModifiedSync().compareTo(a.lastModifiedSync()));
      return fs.isEmpty ? null : fs.first.path;
    } catch (_) {
      return null;
    }
  }

  Future<void> openFile(String? path, String mime) async {
    await openFileStatic(context, path, mime);
  }

  Future<void> openReview() async {
    final list = processedItems.where((e) => e.needsReview).toList();
    if (list.isEmpty) return;
    final changed = await Navigator.push<bool>(
      context,
      MaterialPageRoute(builder: (_) => ReviewPage(items: list, allItems: processedItems)),
    );
    if (changed == true) {
      await _refresh();
    } else if (mounted) {
      setState(() {});
    }
  }

  Future<void> editReceipt(ReceiptItem item) async {
    final changed = await Navigator.push<bool>(
      context,
      MaterialPageRoute(
        builder: (_) => ReviewPage(items: [item], allItems: processedItems, single: true),
      ),
    );
    if (changed == true) {
      await _refresh();
    } else if (mounted) {
      setState(() {});
    }
  }

  void useWhatsApp() {
    saveFolder(null);
    setState(() => customDir = null);
  }

  Future<void> transferOutputs() async {
    final paths = <String>[
      if (latestOutput('.xlsx') != null) latestOutput('.xlsx')!,
      if (latestOutput('.pdf') != null) latestOutput('.pdf')!,
      ...extraPdfPaths,
    ];
    final moved = await transferFilesViaPicker(context, paths.toSet().toList());
    if (moved == null || !mounted) return;
    setState(() {
      if (moved[standardXlsxPath] != null) standardXlsxPath = moved[standardXlsxPath];
      if (moved[generatedPdfPath] != null) generatedPdfPath = moved[generatedPdfPath];
      extraPdfPaths = extraPdfPaths.map((p) => moved[p] ?? p).toList();
    });
  }

  String _countLine() {
    final counted = processedItems
        .where((e) => !e.userExcluded && !e.ignoredDoc && e.ign != 'drop' && e.amount > 0)
        .toList();
    final total = counted.fold<double>(0, (a, e) => a + e.amount);
    final named = processedItems.where((e) => e.tid != null && e.orderNum != null && !e.userExcluded).length;
    final rv = processedItems.where((e) => e.needsReview).length;
    return '${processedItems.length} صورة، $named اتسمّى، إجمالي ${total.toStringAsFixed(2)}'
        '${missingTerminals.isNotEmpty ? ' (ناقص ${missingTerminals.length} ماكينة)' : ''}'
        '${rv > 0 ? ' ($rv للمراجعة)' : ''}';
  }

  Future<void> _markRestOk() async {
    final list = processedItems.where((e) => e.needsReview).toList();
    if (list.isEmpty) return;
    for (final it in list) {
      if (it.dupSuspect) {
        keepAsExtraBatch(it, processedItems);
      } else if (it.dateOdd) {
        it.dateFix = it.date;
        it.dateOdd = false;
        saveFix(it);
      } else if (it.amount > 0 && it.row != null) {
        it.ok = true;
        it.alt = null;
        it.unusual = null;
        it.userExcluded = false;
        it.ignoredDoc = false;
        it.reprintSummary = false;
        it.missingBatchClosed = false;
        if (it.how != 'extra-batch') it.how = 'manual';
        saveFix(it);
      }
    }
    await _refresh();
  }

  Future<void> shareOneReceipt(ReceiptItem item) async {
    if (!File(item.path).existsSync()) return;
    await Share.shareXFiles([XFile(item.path)],
        text: item.orderNum != null ? 'ماكينة ${item.orderNum}' : baseName(item.path));
  }

  Future<void> rebuildFiles() async {
    setState(() => status = 'جاري توليد الملفات وتحديث الشيتات...');
    final dateCounts = <String, int>{};
    for (final r in processedItems) {
      if (r.date != null) {
        dateCounts[r.date!] = (dateCounts[r.date!] ?? 0) + 1;
      }
    }
    String reportDate = "اليوم";
    if (dateCounts.isNotEmpty) {
      reportDate = dateCounts.entries.reduce((a, b) => a.value > b.value ? a : b).key;
    }

    processedItems.sort((a, b) {
      final c = (a.orderNum ?? 99).compareTo(b.orderNum ?? 99);
      return c != 0 ? c : (a.suffix ?? 1).compareTo(b.suffix ?? 1);
    });
    final mainItems = processedItems
        .where((r) => r.counts && (r.date == null || r.date == reportDate))
        .toList();
    checkUnusual(mainItems, reportDate);
    final otherByDate = <String, List<ReceiptItem>>{};
    for (final r in processedItems) {
      if (r.counts && r.date != null && r.date != reportDate) {
        otherByDate.putIfAbsent(r.date!, () => <ReceiptItem>[]).add(r);
      }
    }

    var pageNo = 0;
    Future<bool> buildPdf(List<ReceiptItem> items, File out) async {
      final pdf = pw.Document();
      var pages = 0;
      for (final r in items) {
        pageNo++;
        if (!File(r.path).existsSync()) continue;
        var srcs = <String>[];
        try {
          srcs = await scanReceiptAsync(r.path);
        } catch (_) {}
        if (srcs.isEmpty) srcs = <String>[r.path];
        for (final p in srcs) {
          try {
            final imgWidget = pw.MemoryImage(File(p).readAsBytesSync());
            pdf.addPage(
              pw.Page(
                pageFormat: PdfPageFormat.a4,
                margin: const pw.EdgeInsets.all(20),
                build: (pw.Context ctx) =>
                    pw.Center(child: pw.Image(imgWidget, fit: pw.BoxFit.contain)),
              ),
            );
            pages++;
          } catch (_) {}
        }
      }
      if (pages == 0) return false;
      await out.writeAsBytes(await pdf.save());
      return true;
    }

    final pdfOutFile = File('$outputDir/نتيجة $reportDate.pdf');
    await buildPdf(mainItems, pdfOutFile);
    generatedPdfPath = pdfOutFile.path;
    extraPdfPaths = [];
    for (final e in otherByDate.entries) {
      final f = File('$outputDir/نتيجة ${e.key}.pdf');
      if (await buildPdf(e.value, f)) extraPdfPaths.add(f.path);
    }

    final unitTotals = <int, double>{};
    for (final r in rowNames.keys) {
      unitTotals[r] = 0.0;
    }
    for (final item in mainItems) {
      if (item.row != null && item.amount > 0) {
        unitTotals[item.row!] = (unitTotals[item.row!] ?? 0.0) + item.amount;
      }
    }

    var xlsxLine = '• تم توليد الإكسيل: نتيجة $reportDate.xlsx';
    standardXlsxPath = null;
    final tpl = await loadTemplateBytes();
    if (tpl != null) {
      try {
        final unitRows = rowNames.keys.toList()..sort();
        final outBytes = fillTemplate(tpl, [
          for (final r in unitRows)
            _UnitOut(
              rowNames[r] ?? 'وحدة $r',
              terminals.values.where((t) => t.excelRow == r).length,
              unitTotals[r] ?? 0.0,
            ),
        ]);
        final stdFile = File('$outputDir/نتيجة $reportDate.xlsx');
        stdFile.createSync(recursive: true);
        stdFile.writeAsBytesSync(outBytes);
        standardXlsxPath = stdFile.path;
      } catch (e) {
        xlsxLine = '⚠️ فشل ملء القالب: $e';
      }
    }

    final ignore = <int>{};
    try {
      for (final m in RegExp(r'\d+').allMatches(File('$baseDir/ignore.txt').readAsStringSync())) {
        ignore.add(int.parse(m.group(0)!));
      }
    } catch (_) {}
    final foundOrders =
        mainItems.where((e) => e.amount > 0).map((e) => e.orderNum).whereType<int>().toSet();
    missingTerminals = [
      for (final i in machineNums())
        if (!foundOrders.contains(i) && !ignore.contains(i)) i
    ];
    writeReport(reportDate);

    setState(() => status = 'اكتمل!\n$xlsxLine\n• تم توليد ملف الـ PDF المجمع.');
  }

  String subtitleOf(ReceiptItem item) {
    final sfx = item.suffix != null ? ' (باتش ${item.suffix})' : '';
    final bt = item.batches.length > 1
        ? '\n= ${item.batches.map((b) => b.toStringAsFixed(2)).join(' + ')}'
        : '';
    final alt = item.alt != null ? '\n⚠️ قراءة تانية: ${item.alt!.toStringAsFixed(2)}' : '';
    final txnStr = item.txn != null ? ' | المعاملات: ${item.txn}' : '';
    final extra = StringBuffer();
    if (item.reprintSummary) {
      extra.write('\n🚨 تحذير: الإيصال ده REPRINT أو SUMMARY!');
    } else if (item.missingBatchClosed && !item.ignoredDoc) {
      extra.write('\n⚠️ مش مكتوب فيه BATCH CLOSED SUCCESSFULLY');
    }
    if (item.ignoredDoc) {
      extra.write('\n📄 اتجاهلته: شكله إيصال دفع/مستند');
    }
    if (item.dateOdd) {
      extra.write('\n📅 تاريخه ${item.date} والأغلب ${item.dateMain}');
    }
    final unu = item.unusual != null ? '\n📊 ${item.unusual}' : '';
    final dup = item.dupSuspect ? '\n⚠️ مشكوك إنه مكرر' : (item.userExcluded ? '\n🚫 مستبعد' : '');
    return 'TID: ${item.tid ?? 'غير محدد'}$sfx | المبلغ: ${item.amount.toStringAsFixed(2)} ج.م$txnStr$bt$alt$dup$unu$extra';
  }

  String warningsText() => warnings.map((w) => '• $w').join('\n');

  void writeReport(String date) {
    try {
      final line = '=' * 50;
      final sb = StringBuffer('تقرير المراجعة - $date\n$line\n');
      var total = 0.0;
      for (final r in processedItems) {
        if (r.amount > 0 && r.row != null && r.counts) total += r.amount;
        sb.writeln('${baseName(r.path)}  ${r.unitName ?? '??'}  ${r.amount.toStringAsFixed(2)}');
      }
      sb.writeln(line);
      sb.writeln('الإجمالي العام: ${total.toStringAsFixed(2)} ج.م');
      File('$outputDir/review_report.txt')
        ..createSync(recursive: true)
        ..writeAsStringSync(sb.toString());
    } catch (_) {}
  }

  Future<void> runFullPipeline({List<String>? manualPaths, bool auto = false}) async {
    if (manualPaths == null) _doneSig = _scanSig();
    if (!auto && !chainRunning) {
      try {
        await Permission.notification.request();
      } catch (_) {}
      await askBatteryExempt();
    }
    try {
      await _chan.invokeMethod('startKeepAlive', {'title': 'A2R POS', 'text': 'جاري فحص الإيصالات...'});
      await _chan.invokeMethod('keepScreenOn', {'on': true});
    } catch (_) {}
    final sw = Stopwatch()..start();
    try {
      await _runPipelineInner(manualPaths: manualPaths);
    } finally {
      sw.stop();
      if (mounted && sw.elapsed.inSeconds >= 1 && _isFinalScanStatus(status)) {
        setState(() => status = '$status\n⏱ الوقت المستغرق: ${fmtElapsed(sw.elapsed)}');
      }
      if (autoScan) {
        await _watchService(true);
      } else if (!chainRunning) {
        try {
          await _chan.invokeMethod('stopKeepAlive');
          await _chan.invokeMethod('keepScreenOn', {'on': false});
        } catch (_) {}
      }
    }
  }

  Future<void> _runPipelineInner({List<String>? manualPaths}) async {
    final perm = await Permission.manageExternalStorage.request();
    if (!perm.isGranted) return;
    await loadSettingsFromExcel();

    setState(() {
      busy = true;
      stopReq = false;
      progress = 0;
      processedItems.clear();
      missingTerminals.clear();
      warnings.clear();
      excludedPay.clear();
      status = 'جاري البحث عن مجلدات الصور...';
    });

    final dirs = customDir != null ? [Directory(customDir!)] : findWhatsAppImageDirs();
    if (manualPaths != null) dirs.clear();

    final now = DateTime.now();
    final cutoff = days < 0
        ? DateTime(2000)
        : DateTime(now.year, now.month, now.day).subtract(Duration(days: days));
    final todo = <File>[];
    for (final d in dirs) {
      try {
        for (final e in d.listSync().whereType<File>()) {
          final lower = e.path.toLowerCase();
          if (!(lower.endsWith('.jpg') || lower.endsWith('.jpeg') || lower.endsWith('.png') || lower.endsWith('.webp'))) continue;
          if (baseName(e.path).startsWith('.') || baseName(e.path).startsWith('dup_')) continue;
          if (fileDate(e).isBefore(cutoff)) continue;
          todo.add(e);
        }
      } catch (_) {}
    }
    todo.sort((a, b) => baseName(a.path).compareTo(baseName(b.path)));
    if (manualPaths != null) todo.addAll(manualPaths.map((p) => File(p)));

    if (todo.isEmpty) {
      setState(() {
        busy = false;
        status = 'لا توجد صور في المجلد بالفترة المحددة.';
      });
      return;
    }

    final recognizer = TextRecognizer(script: TextRecognitionScript.latin);
    Directory(rawImagesDir).createSync(recursive: true);
    Directory(outputDir).createSync(recursive: true);
    final fixes = loadFixes();
    final rawReceipts = <ReceiptItem>[];
    final ignoredItems = <ReceiptItem>[];

    for (var i = 0; i < todo.length; i++) {
      if (stopReq) break;
      final f = todo[i];
      try {
        var workPath = f.path;
        var res = await recognizer.processImage(InputImage.fromFilePath(workPath));
        final fkey = '${baseName(f.path)}:${f.lengthSync()}';
        final fxm = fixes[fkey];
        final ign = (fxm is Map && fxm['ign'] is String) ? fxm['ign'] as String : null;
        if (ign != 'keep' && _isIdSlip(res)) continue;

        var isReceipt = receiptScore(res.text) >= 2 || findTerminalIn([res.text]) != null;
        if (ign == 'keep') isReceipt = true;
        var skipPay = false;
        if (isReceipt && ign != 'keep' && classifyReceipt(res.text) == 'payment') {
          isReceipt = false;
          skipPay = true;
        }

        if (isReceipt) {
          final a = analyzeLines(toLines(res));
          double? alt;
          bool hasBatchClosed = RegExp(r'BATCH\s*CLOS', caseSensitive: false).hasMatch(res.text.toUpperCase());
          bool hasReprint = RegExp(r'REPRINT|SUMMARY|DUPLICATE', caseSensitive: false).hasMatch(res.text.toUpperCase());

          final needSecond = !a.ok || a.tid == null || a.row == null || a.val <= 0 || !hasBatchClosed || hasReprint;
          if (needSecond) {
            try {
              final ep = '$tmpDir/enh.jpg';
              await enhanceImageAsync(workPath, ep);
              final res2 = await recognizer.processImage(InputImage.fromFilePath(ep));
              final a2 = analyzeLines(toLines(res2));
              alt = mergeSecondRead(a, a2);
            } catch (_) {}
          }

          final inRaw = f.path.startsWith('$rawImagesDir/');
          final dst = inRaw ? f.path : '$rawImagesDir/wa_${baseName(f.path)}';
          File localFile = File(dst);
          if (!inRaw && !localFile.existsSync()) {
            await fixImageAsync(f.path, dst, 0);
          }

          final item = ReceiptItem(
            path: localFile.path,
            tid: a.tid,
            orderNum: a.tid != null ? terminals[a.tid]!.orderNum : null,
            row: a.row,
            unitName: a.unitName,
            amount: a.val,
            date: a.date,
            key: fkey,
            ok: a.ok,
            how: a.how,
            batches: a.batches,
            alt: alt,
            reprintSummary: hasReprint,
            missingBatchClosed: !hasBatchClosed,
          );
          item.txn = a.txn;
          item.cands = a.cands;
          rawReceipts.add(item);
          processedItems.add(item);
        } else if (skipPay) {
          excludedPay.add(baseName(f.path));
        }
      } catch (e) {
        warnings.add('${baseName(f.path)}: فشلت القراءة');
      }

      if (!mounted) return;
      setState(() {
        progress = ((i + 1) / todo.length) * 0.5;
        status = 'فحص OCR: ${i + 1} من ${todo.length}';
      });
    }

    await recognizer.close();
    processedItems.clear();
    processedItems.addAll(rawReceipts);
    processedItems.addAll(ignoredItems);
    await rebuildFiles();

    if (!mounted) return;
    setState(() {
      busy = false;
      progress = 1.0;
    });
  }

  @override
  Widget build(BuildContext context) {
    AppSettings.of(context);
    final xlsx = latestOutput('.xlsx');
    final pdf = latestOutput('.pdf');
    final reviewCount = processedItems.where((e) => e.needsReview).length;
    _syncBadge(homeBadge, reviewCount);
    _publishClosing(processedItems);
    homeBusyG = busy;
    return Scaffold(
      appBar: AppBar(
        toolbarHeight: 54,
        titleSpacing: 12,
        title: const AppTitle(),
        actions: [
          const HeaderButtons(),
          _moreMenu(context, busy: busy, onLog: () => shareDebugLog(context)),
        ],
      ),
      bottomNavigationBar: _scanBar(
        busy: busy,
        stopReq: stopReq,
        onStop: () => setState(() {
          chainAbort = true;
          stopReq = true;
        }),
        thisLabel: 'فحص التقفيل',
        onScanThis: () => runFullPipeline(),
        progress: progress,
      ),
      body: ListView(
        padding: const EdgeInsets.fromLTRB(8, 6, 8, 8),
        children: [
          _settingsCard([
            _switchRow('فحص تلقائي (بيراقب الإيصالات الجديدة ويفحصها لوحده)', autoScan, (v) => _setAuto(v)),
            const Divider(height: 4),
            InkWell(
              onTap: () => setState(() => settingsOpen = !settingsOpen),
              child: Padding(
                padding: const EdgeInsets.symmetric(vertical: 8),
                child: Row(
                  children: [
                    const Icon(Icons.tune, size: 18),
                    const SizedBox(width: 8),
                    Expanded(
                      child: Text(
                          'إعدادات المصدر • ${_periodLabel()} • ${customDir == null ? 'واتساب' : baseName(customDir!)}',
                          maxLines: 1,
                          overflow: TextOverflow.ellipsis,
                          style: const TextStyle(fontSize: 13)),
                    ),
                    Icon(settingsOpen ? Icons.expand_less : Icons.expand_more),
                  ],
                ),
              ),
            ),
            if (settingsOpen) ...[
              _periodRow(days, busy, (d) {
                setState(() => days = d);
                sharedDays.value = d;
              }),
              const Divider(height: 10),
              _pathRow(Icons.folder_open, 'المصدر', customDir ?? 'واتساب (تلقائي)',
                  busy ? null : pickFolder,
                  onReset: customDir != null && !busy ? useWhatsApp : null),
              _pathRow(Icons.drive_file_move_outline, 'الإخراج',
                  customOut ?? '$defaultOutputDir (افتراضي)', busy ? null : pickOutFolder,
                  onReset: customOut != null && !busy ? useDefaultOut : null,
                  actions: [
                    _pathAct(Icons.open_in_new, 'فتح فولدر الإخراج', 'Open output folder',
                        () => openFolder(outputDir)),
                    _pathAct(Icons.drive_file_move, 'نقل الملفات لفولدر', 'Move files to a folder',
                        busy ? null : transferOutputs),
                  ]),
              _switchRow('فحص الصور المقلوبة (أبطأ)', autoRotate,
                  busy ? null : (v) => setState(() => autoRotate = v)),
            ],
          ]),
          _runAllCard(),
          _statusBox(status),
          const SizedBox(height: 6),
          _statsRow(),
          const SizedBox(height: 6),
          _actionRow([
            _quickBtnG(context, Icons.table_chart, 'إكسيل',
                xlsx == null ? null : () => openFile(xlsx, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')),
            _quickBtnG(context, Icons.picture_as_pdf, 'PDF',
                pdf == null ? null : () => openFile(pdf, 'application/pdf')),
            _quickBtnG(context, Icons.chat, 'واتساب', busy ? null : sendToWhatsApp),
            _quickBtnG(context, Icons.fact_check, 'مراجعة ($reviewCount)', openReview),
            _quickBtnG(context, Icons.add_photo_alternate, 'يدوي', busy ? null : pickManualImages),
          ]),
          if (processedItems.isNotEmpty) ...[
            const SizedBox(height: 4),
            _filterRow(reviewOnly, processedItems.length, reviewCount, (v) => setState(() => reviewOnly = v)),
            _reviewTools(context,
                hint: reviewCount > 0 ? '💡 اسحب إيصال المراجعة: صح ✓ أو غلط ✗' : '',
                count: reviewCount,
                reviewOnly: reviewOnly,
                busy: busy,
                onRestOk: _markRestOk),
          ],
          const SizedBox(height: 10),
          for (final item in processedItems.where((e) => !reviewOnly || e.needsReview))
            _swipeWrap(item, Card(
              margin: const EdgeInsets.symmetric(vertical: 4),
              shape: RoundedRectangleBorder(
                borderRadius: BorderRadius.circular(12),
                side: BorderSide(
                  color: item.reprintSummary 
                      ? Colors.red 
                      : (item.missingBatchClosed ? Colors.orange : _unitColor(item.row)), 
                  width: item.reprintSummary ? 3.0 : (item.missingBatchClosed ? 2.5 : 1.5)
                ),
              ),
              child: ListTile(
                onTap: () => editReceipt(item),
                leading: Image.file(
                  File(item.path),
                  width: 48,
                  height: 48,
                  fit: BoxFit.cover,
                ),
                title: Text(
                  'ماكينة ${item.orderNum ?? '؟'} - ${item.unitName ?? 'وحدة غير معروفة'}',
                  style: const TextStyle(fontWeight: FontWeight.bold),
                ),
                subtitle: Text(subtitleOf(item)),
                trailing: IconButton(
                  icon: Icon(Icons.edit, color: Theme.of(context).colorScheme.primary),
                  onPressed: () => editReceipt(item),
                ),
              ),
            )),
          const SizedBox(height: 64),
        ],
      ),
    );
  }
}

// ───────────────────── متصفح الفولدرات ─────────────────────
class FolderPicker extends StatefulWidget {
  final String start;
  const FolderPicker({super.key, required this.start});

  @override
  State<FolderPicker> createState() => _FolderPickerState();
}

class _FolderPickerState extends State<FolderPicker> {
  late String path = widget.start;
  List<Directory> dirs = [];

  @override
  void initState() {
    super.initState();
    load();
  }

  void load() {
    try {
      dirs = Directory(path)
          .listSync()
          .whereType<Directory>()
          .where((d) => !baseName(d.path).startsWith('.'))
          .toList()
        ..sort((a, b) => baseName(a.path).toLowerCase().compareTo(baseName(b.path).toLowerCase()));
    } catch (_) {
      dirs = [];
    }
  }

  void go(String p) => setState(() {
        path = p;
        load();
      });

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: const Text('اختيار فولدر')),
      body: SafeArea(
        child: Column(
          children: [
            Padding(
              padding: const EdgeInsets.all(12),
              child: Text(path, textDirection: TextDirection.ltr),
            ),
            Expanded(
              child: ListView(
                children: [
                  if (path.length > root.length)
                    ListTile(
                      leading: const Icon(Icons.arrow_upward),
                      title: const Text('..'),
                      onTap: () => go(Directory(path).parent.path),
                    ),
                  for (final d in dirs)
                    ListTile(
                      leading: const Icon(Icons.folder),
                      title: Text(baseName(d.path), textDirection: TextDirection.ltr),
                      onTap: () => go(d.path),
                    ),
                ],
              ),
            ),
            Padding(
              padding: const EdgeInsets.all(12),
              child: FilledButton(
                onPressed: () => Navigator.pop(context, path),
                child: const Text('اختيار هذا الفولدر'),
              ),
            ),
          ],
        ),
      ),
    );
  }
}

// ───────────────────── شاشة المراجعة ─────────────────────
class ZoomableImage extends StatefulWidget {
  final String path;
  const ZoomableImage({super.key, required this.path});

  @override
  State<ZoomableImage> createState() => _ZoomableImageState();
}

class _ZoomableImageState extends State<ZoomableImage> {
  final _tc = TransformationController();
  Offset _tap = Offset.zero;

  @override
  void dispose() {
    _tc.dispose();
    super.dispose();
  }

  void _toggle() {
    if (_tc.value.getMaxScaleOnAxis() > 1.05) {
      _tc.value = Matrix4.identity();
    } else {
      const s = 3.0;
      final tx = -_tap.dx * (s - 1);
      final ty = -_tap.dy * (s - 1);
      _tc.value = Matrix4(s, 0, 0, 0, 0, s, 0, 0, 0, 0, 1, 0, tx, ty, 0, 1);
    }
  }

  @override
  Widget build(BuildContext context) {
    return GestureDetector(
      onDoubleTapDown: (d) => _tap = d.localPosition,
      onDoubleTap: _toggle,
      child: InteractiveViewer(
        transformationController: _tc,
        minScale: 1,
        maxScale: 12,
        child: Center(
          child: Image.file(
            File(widget.path),
            fit: BoxFit.contain,
            filterQuality: FilterQuality.high,
          ),
        ),
      ),
    );
  }
}

class ReceiptZoomPage extends StatelessWidget {
  final String path;
  final String title;
  const ReceiptZoomPage({super.key, required this.path, required this.title});

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      backgroundColor: Colors.black,
      appBar: AppBar(
        title: Text(title, style: const TextStyle(fontSize: 15)),
        backgroundColor: Colors.black,
        foregroundColor: Colors.white,
      ),
      body: SafeArea(child: ZoomableImage(path: path)),
    );
  }
}

class ReviewPage extends StatefulWidget {
  final List<ReceiptItem> items;
  final List<ReceiptItem>? allItems;
  final bool single;
  const ReviewPage({super.key, required this.items, this.allItems, this.single = false});

  @override
  State<ReviewPage> createState() => _ReviewPageState();
}

class _ReviewPageState extends State<ReviewPage> {
  int idx = 0;
  bool changed = false;
  int rev = 0;
  final done = <ReceiptItem>{};
  late TextEditingController amt;
  int? order;

  ReceiptItem get item => widget.items[idx];

  @override
  void initState() {
    super.initState();
    amt = TextEditingController();
    load();
  }

  @override
  void dispose() {
    amt.dispose();
    super.dispose();
  }

  void load() {
    amt.text = item.amount > 0 ? item.amount.toStringAsFixed(2) : '';
    order = item.orderNum;
  }

  void go(int i) {
    if (i < 0 || i >= widget.items.length) return;
    setState(() {
      idx = i;
      load();
    });
  }

  void save({bool next = true}) {
    final v = double.tryParse(amt.text.replaceAll(',', '.'));
    if (v == null || v <= 0) return;
    final it = item;
    it.amount = round2(v);
    it.batches = [it.amount];
    it.ok = true;
    it.alt = null;
    it.how = 'manual';
    it.unusual = null;
    it.reprintSummary = false;
    it.missingBatchClosed = false;
    if (order != null) {
      it.orderNum = order;
      it.tid = numToTid[order];
      it.row = terminals[it.tid]!.excelRow;
      it.unitName = terminals[it.tid]!.unitName;
    }
    saveFix(it);
    changed = true;
    done.add(it);
    if (widget.single) {
      Navigator.pop(context, true);
    } else if (next && idx < widget.items.length - 1) {
      go(idx + 1);
    } else {
      setState(() {});
    }
  }

  @override
  Widget build(BuildContext context) {
    AppSettings.of(context);
    final it = item;
    return Scaffold(
      appBar: AppBar(title: Text('مراجعة ${idx + 1} / ${widget.items.length}')),
      body: SafeArea(
        child: Column(
          children: [
            Expanded(flex: 5, child: ZoomableImage(key: ValueKey('${it.path}#$rev'), path: it.path)),
            Expanded(
              flex: 4,
              child: ListView(
                padding: const EdgeInsets.all(12),
                children: [
                  if (it.cands.isNotEmpty) ...[
                    const Text('مبالغ مقترحة من الإيصال:', style: TextStyle(fontSize: 12)),
                    Wrap(
                      spacing: 6,
                      children: [
                        for (final v in it.cands)
                          ActionChip(
                            label: Text(v.toStringAsFixed(2)),
                            onPressed: () => setState(() => amt.text = v.toStringAsFixed(2)),
                          ),
                      ],
                    ),
                    const SizedBox(height: 8),
                  ],
                  TextField(
                    controller: amt,
                    keyboardType: const TextInputType.numberWithOptions(decimal: true),
                    decoration: const InputDecoration(labelText: 'المبلغ الإجمالي', border: OutlineInputBorder()),
                  ),
                  const SizedBox(height: 8),
                  DropdownButtonFormField<int>(
                    value: order,
                    isExpanded: true,
                    hint: const Text('اختر الماكينة'),
                    items: [
                      for (final i in machineNums())
                        DropdownMenuItem(
                          value: i,
                          child: Text(machineLabel(i)),
                        ),
                    ],
                    onChanged: (v) => setState(() => order = v),
                  ),
                  const SizedBox(height: 8),
                  FilledButton(
                    onPressed: () => save(),
                    child: Text(idx < widget.items.length - 1 ? 'حفظ والتالي' : 'حفظ'),
                  ),
                ],
              ),
            ),
          ],
        ),
      ),
    );
  }
}

// ───────────────────── تاب إيصالات الدفع والملفات ─────────────────────
String? loadPdfDir() => _readCfg()['pdfDir'] as String?;
List<Directory> findWhatsAppDocDirs() {
  final out = <Directory>[];
  for (final d in findWhatsAppImageDirs()) {
    final p = d.path;
    if (p.endsWith('WhatsApp Images')) {
      final dd = Directory('${p.substring(0, p.length - 'WhatsApp Images'.length)}WhatsApp Documents');
      if (dd.existsSync() && !out.any((o) => o.path == dd.path)) out.add(dd);
    }
  }
  return out;
}

void savePdfDir(String? f) {
  final m = _readCfg();
  m['pdfDir'] = f;
  _writeCfg(m);
}

String machineNewName(TerminalInfo info, String day, String month) {
  final same = terminals.values.where((t) => t.excelRow == info.excelRow).toList()
    ..sort((a, b) => a.orderNum.compareTo(b.orderNum));
  final unit = info.unitName.trim().replaceAll(RegExp(r'\s+'), '.');
  final k = same.indexWhere((t) => t.orderNum == info.orderNum) + 1;
  return [
    info.orderNum.toString(),
    unit,
    if (same.length > 1) '$k',
    'يوم',
    day,
    month,
  ].join('.');
}

final RegExp _dateRe = RegExp(r'(20\d{2})[^\d\n]{1,3}(\d{2})[^\d\n]{1,3}(\d{2})');

String? _dayKey(String text) {
  for (final m in _dateRe.allMatches(text)) {
    final mo = int.parse(m.group(2)!), d = int.parse(m.group(3)!);
    if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31) return '$d.$mo';
  }
  return null;
}

String? _anchorOf(String t) {
  final m = RegExp(r'R{1,2}N\W{0,3}((?:\d\s?){9,12})', caseSensitive: false).firstMatch(t) ??
      RegExp(r'(?<!\d)(63(?:\d\s?){7,10})(?!\d)').firstMatch(t);
  if (m == null) return null;
  final d = m.group(1)!.replaceAll(RegExp(r'\s'), '');
  if (d.length < 9) return null;
  if (d.length == 12 && d.startsWith('63')) return d;
  return 'P';
}

const double _stripScale = 1.5;

List<int> _makeStrips(String src, String prefix) {
  final im = img.decodeImage(File(src).readAsBytesSync());
  if (im == null) return [];
  final w = im.width, h = im.height;
  const n = 4;
  final sw = w * 0.40;
  final step = (w - sw) / (n - 1);
  final out = <int>[w, h];
  for (var i = 0; i < n; i++) {
    final x0 = (i * step).round();
    final x1 = math.min(w, (x0 + sw).round());
    var c = img.copyCrop(im, x: x0, y: 0, width: x1 - x0, height: h);
    c = img.copyResize(c,
        width: (c.width * _stripScale).round(), interpolation: img.Interpolation.linear);
    File('${prefix}_$i.jpg').writeAsBytesSync(img.encodeJpg(c, quality: 90));
    out.add(x0);
  }
  return out;
}

class Anchor {
  String? rrn;
  String? day;
  final double l, r, cy;
  Anchor(this.rrn, this.day, this.l, this.r, this.cy);
}

class MCollector {
  final List<List<Anchor>> pages = [];
  int slips = 0;
  final StringBuffer text = StringBuffer();
  double pageH = 3000;

  void add(Anchor n) {
    final list = pages.last;
    for (final a in list) {
      if (a.l < n.r && n.l < a.r && (a.cy - n.cy).abs() < pageH * 0.012) {
        a.rrn ??= n.rrn;
        a.day ??= n.day;
        return;
      }
    }
    list.add(n);
  }

  MScan finish() {
    final r = MScan();
    final seen = <String>{};
    final days = <String, int>{};
    final partialDays = <String>[];
    for (final a in pages.expand((p) => p)) {
      if (a.rrn != null) {
        if (seen.add(a.rrn!)) {
          r.full++;
          r.rrnDay[a.rrn!] = a.day;
          if (a.day != null) days[a.day!] = (days[a.day!] ?? 0) + 1;
        }
      } else {
        r.partial++;
        if (a.day != null) {
          partialDays.add(a.day!);
          days[a.day!] = (days[a.day!] ?? 0) + 1;
        }
      }
    }
    final all = text.toString();
    r.nonPayment = RegExp(
            r'GRAND\s*TOTAL|TOTAL\s*NET\s*AMOUNT|BATCH\s*CLOS|SUMMARY|REPRINT|DUPLICATE|SUCCESSF',
            caseSensitive: false)
        .hasMatch(all);

    final tc = <String, int>{};
    for (final m in RegExp(r'(?<!\d)(\d{8})(?!\d)').allMatches(all)) {
      var c = m.group(1)!;
      c = tidFix[c] ?? c;
      if (terminals.containsKey(c)) tc[c] = (tc[c] ?? 0) + 1;
    }
    r.tid = tc.isNotEmpty
        ? tc.entries.reduce((a, b) => a.value >= b.value ? a : b).key
        : (findTerminalIn([all]) ?? discoverTerminal([all]));

    if (days.isNotEmpty) {
      final best = days.entries.reduce((a, b) => a.value >= b.value ? a : b).key;
      final p = best.split('.');
      r.day = p[0];
      r.month = p[1];
    }
    return r;
  }
}

class MScan {
  int full = 0, partial = 0;
  final Map<String, String?> rrnDay = {};
  bool nonPayment = false;
  String? tid, day, month;
  final List<String> otherDay = [];
  int? seqMin, seqMax;
  final List<int> seqMissing = [];

  int get count => full + partial;
  int? get rangeLen => seqMin == null ? null : seqMax! - seqMin! + 1;
  String get rangeText =>
      seqMin == null ? '' : ' • التسلسل ${seqMin.toString().padLeft(6, '0')}-${seqMax.toString().padLeft(6, '0')} ($rangeLen)';
}

void _collect(RecognizedText res, double x0, double scale, MCollector col) {
  col.text.writeln(res.text);
  final lines = toLines(res);
  final dates = lines.where((l) => _dayKey(l.text) != null).toList();
  for (final l in lines) {
    final a = _anchorOf(l.text);
    if (a == null) continue;
    var day = _dayKey(l.text);
    if (day == null) {
      OcrLine? best;
      var bd = double.infinity;
      for (final d in dates) {
        final dy = (d.cy - l.cy).abs();
        if (dy > l.h * 5) continue;
        if (!(d.left < l.right + 60 && l.left < d.right + 60)) continue;
        if (dy < bd) {
          bd = dy;
          best = d;
        }
      }
      if (best != null) day = _dayKey(best.text);
    }
    col.add(Anchor(a == 'P' ? null : a, day, x0 + l.left / scale, x0 + l.right / scale,
        l.cy / scale));
  }
}

Future<void> ocrPageStrips(
    TextRecognizer rec, String pngPath, String prefix, MCollector col) async {
  final stripsF = Isolate.run(() => _makeStrips(pngPath, prefix))
      .then<List<int>>((v) => v, onError: (_) => <int>[]);
  RecognizedText? whole;
  try {
    whole = await _ocr(rec, pngPath);
  } catch (_) {}
  final info = await stripsF;
  col.pageH = info.length >= 2 ? info[1].toDouble() : 3000;
  col.pages.add(<Anchor>[]);
  if (whole != null && _isIdSlip(whole)) {
    col.slips++;
    return;
  }
  if (whole != null) _collect(whole, 0, 1, col);
  for (var i = 2; i < info.length; i++) {
    final fp = '${prefix}_${i - 2}.jpg';
    try {
      final res = await _ocr(rec, fp);
      _collect(res, info[i].toDouble(), _stripScale, col);
    } catch (_) {}
    try {
      File(fp).deleteSync();
    } catch (_) {}
  }
}

final String mStatePath = '$baseDir/.app_machines.json';
final String defaultMDest = '$baseDir/وحدات';

String loadMOut() => (_readCfg()['mOut'] as String?) ?? defaultMDest;
void saveMOut(String? f) {
  final m = _readCfg();
  m['mOut'] = f;
  _writeCfg(m);
}

class MResult {
  final String pdfName;
  String srcPath = '';
  int size = 0;
  String? openPath;
  int? orderNum;
  String? newName;
  int count = -1;
  int tries = 0;
  String detail = '';
  final List<String> idNotes = [];
  final List<String> cntNotes = [];
  List<String> altTids = [];
  Map<String, double> votes = {};
  String? tidUsed, dayUsed, monthUsed;
  bool idOk = true;
  bool cntOk = true;
  int verdict = 0;
  final TextEditingController ctl = TextEditingController();
  MResult(this.pdfName);

  bool get excluded => verdict < 0;
  bool get pending => count < 0 && !excluded;
  bool get ok => verdict != 0 || (idOk && cntOk);
  String get key => '$pdfName:$size';
  List<String> get notes => [...idNotes, ...cntNotes];

  Map<String, dynamic> toJson() => {
        'pdfName': pdfName,
        'srcPath': srcPath,
        'size': size,
        'openPath': openPath,
        'orderNum': orderNum,
        'tidUsed': tidUsed,
        'newName': newName,
        'count': count,
        'tries': tries,
        'detail': detail,
        'idNotes': idNotes,
        'cntNotes': cntNotes,
        'idOk': idOk,
        'cntOk': cntOk,
        'verdict': verdict,
        'exp': ctl.text,
      };

  static MResult fromJson(Map<String, dynamic> j) {
    final r = MResult((j['pdfName'] ?? '').toString());
    r.srcPath = (j['srcPath'] ?? '').toString();
    r.size = (j['size'] as num?)?.toInt() ?? 0;
    r.openPath = j['openPath'] as String?;
    r.orderNum = (j['orderNum'] as num?)?.toInt();
    r.tidUsed = j['tidUsed'] as String?;
    r.newName = j['newName'] as String?;
    r.count = (j['count'] as num?)?.toInt() ?? -1;
    r.tries = (j['tries'] as num?)?.toInt() ?? 0;
    r.detail = (j['detail'] ?? '').toString();
    r.idNotes.addAll(((j['idNotes'] as List?) ?? const []).map((e) => e.toString()));
    r.cntNotes.addAll(((j['cntNotes'] as List?) ?? const []).map((e) => e.toString()));
    r.idOk = j['idOk'] != false;
    r.cntOk = j['cntOk'] != false;
    r.verdict = (j['verdict'] as num?)?.toInt() ?? 0;
    r.ctl.text = (j['exp'] ?? '').toString();
    return r;
  }
}

List<MResult> loadMState() {
  try {
    final f = File(mStatePath);
    if (!f.existsSync()) return [];
    final l = jsonDecode(f.readAsStringSync());
    if (l is! List) return [];
    return l.whereType().map((e) => MResult.fromJson(Map<String, dynamic>.from(e))).toList();
  } catch (_) {
    return [];
  }
}

void saveMState(List<MResult> rs) {
  try {
    File(mStatePath)
      ..createSync(recursive: true)
      ..writeAsStringSync(jsonEncode(rs.map((r) => r.toJson()).toList()));
  } catch (_) {}
}

class MIdent {
  String? tid, day, month;
}

MIdent _identify(String all) {
  final r = MIdent();
  final tc = <String, int>{};
  for (final m in RegExp(r'(?<!\d)(\d{8})(?!\d)').allMatches(all)) {
    var c = m.group(1)!;
    c = tidFix[c] ?? c;
    if (terminals.containsKey(c)) tc[c] = (tc[c] ?? 0) + 1;
  }
  r.tid = tc.isNotEmpty
      ? tc.entries.reduce((a, b) => a.value >= b.value ? a : b).key
      : (findTerminalIn([all]) ?? discoverTerminal([all]));
  final days = <String, int>{};
  for (final m in _dateRe.allMatches(all)) {
    final k = _dayKey(m.group(0)!);
    if (k != null) days[k] = (days[k] ?? 0) + 1;
  }
  if (days.isNotEmpty) {
    final best = days.entries.reduce((a, b) => a.value >= b.value ? a : b).key;
    final p = best.split('.');
    r.day = p[0];
    r.month = p[1];
  }
  return r;
}

Map<String, double> _tidVotes(String text) {
  final v = <String, double>{};
  final lab = _tidLabeled(text);
  if (lab != null) v[lab] = (v[lab] ?? 0) + 3;
  for (final m in RegExp(r'(?<!\d)(\d{8})(?!\d)').allMatches(text)) {
    final c0 = m.group(1)!;
    final c = tidFix[c0] ?? c0;
    if (terminals.containsKey(c)) v[c] = (v[c] ?? 0) + 1;
  }
  return v;
}

List<MapEntry<String, double>> _rankVotes(Map<String, double> votes) {
  final l = votes.entries.toList()..sort((a, b) => b.value.compareTo(a.value));
  return l;
}

int _cmpM(MResult a, MResult b) {
  final c = (a.orderNum ?? 99).compareTo(b.orderNum ?? 99);
  return c != 0 ? c : (a.newName ?? a.pdfName).compareTo(b.newName ?? b.pdfName);
}

String _l(BuildContext c, String ar, String en) => AppSettings.of(c).isAr ? ar : en;

int? _batchManual(MResult r) => int.tryParse(r.ctl.text.trim());
int? _batchAutoOf(MResult r) => r.orderNum == null ? null : closingTxn.value[r.orderNum!];
int? _batchOf(MResult r) => _batchManual(r) ?? _batchAutoOf(r);

bool? _batchMatch(MResult r) {
  if (r.excluded || r.pending) return null;
  if (r.verdict > 0) return true;
  final b = _batchOf(r);
  if (b == null) return null;
  return b == r.count;
}

String _mSummary(BuildContext c, MResult r) {
  if (r.excluded) return _l(c, 'مستبعد', 'Excluded');
  if (r.pending) return _l(c, 'لسه ما اتعدّش', 'Not counted yet');
  final bt = _batchOf(r);
  if (bt != null) {
    return 'معاملات: ${r.count} • الباتش: $bt';
  }
  return 'معاملات: ${r.count}';
}

class MachinesPage extends StatefulWidget {
  const MachinesPage({super.key});

  @override
  State<MachinesPage> createState() => _MachinesPageState();
}

class _MachinesPageState extends State<MachinesPage> with WidgetsBindingObserver {
  String? pdfDir;
  int days = 0;
  bool skipDone = true;
  bool reviewOnly = false;
  bool busy = false;
  bool stopReq = false;
  bool settingsOpen = false;
  double progressValue = 0;
  String status = 'اضغط "فحص وإعادة تسمية" لفحص إيصالات الدفع.';
  final List<MResult> results = [];
  String destDir = loadMOut();
  bool autoScan = false;
  Timer? _autoTimer;

  @override
  void initState() {
    super.initState();
    pdfDir = loadPdfDir();
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
    setState(() {
      results.clear();
      status = 'تم مسح كل حاجة.';
    });
  }

  void _onClosing() {
    if (mounted) setState(() {});
  }

  Future<void> openPdf(String? path) async {
    await openFileStatic(context, path, 'application/pdf');
  }

  Future<void> pickDir() async {
    final perm = await Permission.manageExternalStorage.request();
    if (!perm.isGranted) return;
    if (!mounted) return;
    final res = await Navigator.push<String>(
      context,
      MaterialPageRoute(builder: (_) => FolderPicker(start: pdfDir ?? root)),
    );
    if (res != null) {
      savePdfDir(res);
      setState(() => pdfDir = res);
    }
  }

  Future<void> pickOut() async {
    final perm = await Permission.manageExternalStorage.request();
    if (!perm.isGranted) return;
    if (!mounted) return;
    final res = await Navigator.push<String>(
      context,
      MaterialPageRoute(builder: (_) => FolderPicker(start: destDir)),
    );
    if (res != null) {
      saveMOut(res);
      setState(() => destDir = res);
    }
  }

  void useDefaultOut() {
    saveMOut(null);
    setState(() => destDir = defaultMDest);
  }

  Future<void> shareAll() async {
    final files = results
        .where((r) => !r.excluded && r.openPath != null && File(r.openPath!).existsSync())
        .map((r) => XFile(r.openPath!, mimeType: 'application/pdf'))
        .toList();
    if (files.isEmpty) return;
    await Share.shareXFiles(files, text: 'إيصالات الدفع');
  }

  void setVerdict(MResult r, int v, {bool toggle = true}) {
    r.verdict = (toggle && r.verdict == v) ? 0 : v;
    saveMState(results);
    setState(() {});
  }

  Future<void> start({bool resumeOnly = false, bool count = false}) async {
    if (busy) return;
    final perm = await Permission.manageExternalStorage.request();
    if (!perm.isGranted) return;
    await _run(resumeOnly, count);
  }

  Future<void> _run(bool resumeOnly, bool count) async {
    await prepareDefaultFiles();
    setState(() {
      busy = true;
      stopReq = false;
      progressValue = 0;
      status = 'جاري البحث عن ملفات PDF...';
    });

    var rec = TextRecognizer(script: TextRecognitionScript.latin);
    try {
      if (!resumeOnly) await _identifyAll(rec);
      if (count) await _countAll(rec);
    } finally {
      await rec.close();
    }
    results.sort(_cmpM);
    saveMState(results);
    if (!mounted) return;
    setState(() {
      busy = false;
      status = 'اكتمل العد: ${results.length} ملف.';
    });
  }

  Future<void> _identifyAll(TextRecognizer rec) async {
    final dirs = pdfDir != null ? [Directory(pdfDir!)] : findWhatsAppDocDirs();
    final pdfs = <File>[];
    for (final d in dirs) {
      try {
        for (final e in d.listSync().whereType<File>()) {
          if (e.path.toLowerCase().endsWith('.pdf') && !_isAppReportName(baseName(e.path))) {
            pdfs.add(e);
          }
        }
      } catch (_) {}
    }
    for (var i = 0; i < pdfs.length; i++) {
      if (stopReq) break;
      final f = pdfs[i];
      final r = MResult(baseName(f.path))
        ..srcPath = f.path
        ..size = f.lengthSync();
      await _identifyOne(r, f, rec);
      r.openPath ??= f.path;
      results.add(r);
    }
  }

  Future<void> _identifyOne(MResult r, File f, TextRecognizer rec) async {
    final bytes = await f.readAsBytes();
    final sb = StringBuffer();
    final votes = <String, double>{};
    var id = MIdent();
    for (var pi = 0; pi < 10; pi++) {
      try {
        await for (final page in Printing.raster(bytes, pages: [pi], dpi: 150)) {
          final png = await page.toPng();
          final tp = File('$mTmpDir/id_$pi.png');
          _ensureMTmp();
          await tp.writeAsBytes(png);
          final res = await _ocr(rec, tp.path);
          sb.writeln(res.text);
          _tidVotes(res.text).forEach((k, v) => votes[k] = (votes[k] ?? 0) + v);
          try { tp.deleteSync(); } catch (_) {}
        }
      } catch (_) { break; }
    }
    final rk = _rankVotes(votes);
    var tid = rk.isNotEmpty ? rk.first.key : id.tid;
    id = _identify(sb.toString());
    if (tid != null && terminals.containsKey(tid) && id.day != null && id.month != null) {
      final info = terminals[tid]!;
      r.orderNum = info.orderNum;
      r.newName = machineNewName(info, id.day!, id.month!);
      final dst = File('$destDir/${r.newName}.pdf');
      await f.copy(dst.path);
      r.openPath = dst.path;
    }
  }

  Future<void> _countAll(TextRecognizer rec) async {
    final todo = results.where((r) => r.pending).toList();
    for (var i = 0; i < todo.length; i++) {
      if (stopReq) break;
      final r = todo[i];
      final src = File(r.openPath ?? r.srcPath);
      if (!src.existsSync()) continue;
      final bytes = await src.readAsBytes();
      final ranges = await Isolate.run(() => _pdfJpegRanges(bytes));
      var total = 0;
      for (var idx = 0; idx < ranges.length; idx++) {
        final fp = File('$mTmpDir/f_$idx.jpg');
        _ensureMTmp();
        await fp.writeAsBytes(Uint8List.sublistView(bytes, ranges[idx][0], ranges[idx][1]));
        final res = await _ocr(rec, fp.path);
        if (!_isIdSlip(res)) {
          final sp = _structCount(res);
          total += (sp.sm3 > 0 ? sp.sm3 : sp.sm);
        }
        try { fp.deleteSync(); } catch (_) {}
      }
      r.count = total;
    }
  }

  @override
  Widget build(BuildContext context) {
    AppSettings.of(context);
    final reviewN = results.where((r) => !r.ok).length;
    _syncBadge(machBadge, reviewN);
    machBusyG = busy;
    return Scaffold(
      appBar: AppBar(
        title: const AppTitle(),
        actions: [
          const HeaderButtons(),
          _moreMenu(context, busy: busy, onLog: () => shareDebugLog(context)),
        ],
      ),
      bottomNavigationBar: _scanBar(
        busy: busy,
        stopReq: stopReq,
        onStop: () => setState(() => stopReq = true),
        thisLabel: 'فحص إيصالات الدفع',
        onScanThis: () => start(count: true),
        progress: progressValue,
      ),
      body: ListView(
        padding: const EdgeInsets.fromLTRB(8, 6, 8, 8),
        children: [
          _settingsCard([
            _switchRow('العد المزدوج ومقارنة الطرق', dualCount, (v) {
              toggleDualCount();
              setState(() {});
            }),
            _pathRow(Icons.folder_open, 'المصدر', pdfDir ?? 'واتساب - مستندات', busy ? null : pickDir),
            _pathRow(Icons.drive_file_move_outline, 'الإخراج', destDir, busy ? null : pickOut),
          ]),
          _statusBox(status),
          const SizedBox(height: 6),
          _actionRow([
            _quickBtnG(context, Icons.chat, 'إرسال الكل', shareAll),
          ]),
          const SizedBox(height: 10),
          for (final r in results.where((r) => !reviewOnly || !r.ok))
            Card(
              margin: const EdgeInsets.symmetric(vertical: 2),
              child: ListTile(
                title: Text(r.newName ?? r.pdfName),
                subtitle: w.Text(_mSummary(context, r)),
                trailing: SizedBox(
                  width: 66,
                  child: TextField(
                    controller: r.ctl,
                    keyboardType: TextInputType.number,
                    textAlign: TextAlign.center,
                    decoration: InputDecoration(
                      hintText: _batchAutoOf(r)?.toString() ?? 'الصح',
                      border: const OutlineInputBorder(),
                    ),
                    onChanged: (_) {
                      saveMState(results);
                      setState(() {});
                    },
                  ),
                ),
              ),
            ),
        ],
      ),
    );
  }
}
