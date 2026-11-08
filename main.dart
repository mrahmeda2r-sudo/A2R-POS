import 'dart:convert';
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

// ط¨ظٹط®طھط§ط± DPI ط¨ط­ظٹط« ط§ظ„ط¶ظ„ط¹ ط§ظ„ط£ط·ظˆظ„ ظ…ط§ ظٹط²ظٹط¯ط´ ط¹ظ† ~2500 ط¨ظƒط³ظ„ (PDF ط§ظ„ظ…ظ…ط³ظˆط­ ط¨ظٹط¨ظ‚ظ‰ ظ…ظ‚ط§ط³ظ‡ ظƒط¨ظٹط± ظˆط¨ظٹط®ظ„ظ‘ظٹ ط§ظ„ط°ط§ظƒط±ط© طھط®ظ„طµ)
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

// طھط§ط±ظٹط® ط¥ظٹطµط§ظ„ ظپظٹ ط§ظ„ظ…ط³طھظ‚ط¨ظ„ = ظ‚ط±ط§ط،ط© ط؛ظ„ط·
bool _plausibleDay(String? d, String? m) {
  if (d == null || m == null) return true;
  final di = int.tryParse(d), mi = int.tryParse(m);
  if (di == null || mi == null) return false;
  final n = DateTime.now();
  final dt = DateTime(n.year, mi, di);
  return !dt.isAfter(DateTime(n.year, n.month, n.day).add(const Duration(days: 1)));
}

bool _pageEnd(Object e) => e.toString().contains('Invalid page index');

// ط³ط¬ظ„ طھط´ط®ظٹطµظٹ ط¨ظٹطھظƒطھط¨ ظپظٹ id_debug.txt (ط¨ظٹط³ط§ط¹ط¯ ظ†ط¹ط±ظپ ظ„ظٹظ‡ ظ…ظ„ظپ ط§طھط®ط·ظ‘ظ‰ ط£ظˆ ط§ظ„طھط·ط¨ظٹظ‚ ط®ط±ط¬)
void _dbg(String m) {
  try {
    final t = DateTime.now().toIso8601String().substring(11, 19);
    File('$baseDir/id_debug.txt').writeAsStringSync('[$t] $m\n', mode: FileMode.append);
  } catch (_) {}
}

// ظ…ظ„ظپ ط§ظ„ظ†طھظٹط¬ط© ط§ظ„ظ„ظٹ ط§ظ„طھط·ط¨ظٹظ‚ ظ†ظپط³ظ‡ ط¨ظٹط·ظ„ظ‘ط¹ظ‡ (ظ†طھظٹط¬ط© 5.11.pdf) ط¨ط³طŒ ظ…ط´ ط£ظٹ ط§ط³ظ… طھط§ظ†ظٹ
bool _isAppReportName(String n) => RegExp(
        r'^(ظ†طھظٹط¬ظ‡|ظ†ط§طھط¬|ط§ظ„ط§ط³ظƒظ†ط¯ط±ظٹظ‡)\s*(\d{1,2}[.\-_]\d{1,2}|ط§ظ„ظٹظˆظ…)(\s*\(\d+\))?\.pdf$')
    .hasMatch(_normAr(n));

// ط£ظٹ ظ…ظ„ظپ ط¥ط®ط±ط§ط¬ ط·ظ„ظ‘ط¹ظ‡ ط§ظ„طھط·ط¨ظٹظ‚ (ط¥ظƒط³ظٹظ„ ط£ظˆ PDF): "ظ†طھظٹط¬ط© 5.11.xlsx" (ظˆط§ظ„ط§ط³ظ… ط§ظ„ظ‚ط¯ظٹظ… ظ„ط³ظ‡ ظ…طھط¹ط±ظ‘ظپ ط¹ظ„ظٹظ‡)
bool _isOutputFileName(String n) => RegExp(
        r'^(ظ†طھظٹط¬ظ‡|ظ†ط§طھط¬|ط§ظ„ط§ط³ظƒظ†ط¯ط±ظٹظ‡)\s*(\d{1,2}[.\-_]\d{1,2}|ط§ظ„ظٹظˆظ…)(\s*\(\d+\))?\.(pdf|xlsx)$')
    .hasMatch(_normAr(n));

// â”€â”€â”€â”€â”€â”€â”€â”€â”€ ط¹ظ†ط§طµط± ظˆط§ط¬ظ‡ط© ظ…ط¯ظ…ط¬ط© (ط¹ظ„ط´ط§ظ† ط§ظ„ط´ط§ط´ط© ظ…ط§ طھط¨ظ‚ط§ط´ ط²ط­ظ…ط©) â”€â”€â”€â”€â”€â”€â”€â”€â”€
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
        const Text('ط§ظ„ظپطھط±ط©', style: TextStyle(fontWeight: FontWeight.bold, fontSize: 13)),
        const SizedBox(width: 10),
        Expanded(
          child: Wrap(
            spacing: 6,
            children: [
              for (final o in const [
                [0, 'ط§ظ„ظٹظˆظ…'],
                [1, 'ظٹظˆظ…ظٹظ†'],
                [2, '3 ط£ظٹط§ظ…'],
                [-1, 'ط§ظ„ظƒظ„'],
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
        const Text('ط¹ط±ط¶', style: TextStyle(fontWeight: FontWeight.bold, fontSize: 13)),
        const SizedBox(width: 10),
        Expanded(
          child: Wrap(
            spacing: 6,
            children: [
              ChoiceChip(
                label: Text('ط§ظ„ظƒظ„ ($all)', style: const TextStyle(fontSize: 13)),
                visualDensity: VisualDensity.compact,
                materialTapTargetSize: MaterialTapTargetSize.shrinkWrap,
                selected: !reviewOnly,
                onSelected: (_) => onPick(false),
              ),
              ChoiceChip(
                label: Text('ظ…ط±ط§ط¬ط¹ط© ($review)', style: const TextStyle(fontSize: 13)),
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

Widget _pathRow(IconData icon, String label, String path, VoidCallback? onPick,
    {VoidCallback? onReset}) {
  final bs = TextButton.styleFrom(
    visualDensity: VisualDensity.compact,
    padding: const EdgeInsets.symmetric(horizontal: 8),
    minimumSize: const Size(0, 32),
    tapTargetSize: MaterialTapTargetSize.shrinkWrap,
  );
  return GestureDetector(
    behavior: HitTestBehavior.opaque,
    onLongPress: onReset, // ط¶ط؛ط·ط© ظ…ط·ظˆظ‘ظ„ط© = ط±ط¬ظˆط¹ ظ„ظ„ط§ظپطھط±ط§ط¶ظٹ
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
        TextButton(onPressed: onPick, style: bs, child: const Text('طھط؛ظٹظٹط±')),
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

// ط±ظ‚ظ… ط§ظ„ط¥طµط¯ط§ط±: ط¨ظٹظٹط¬ظٹ ظ…ظ† GitHub Actions ظˆظ‚طھ ط§ظ„ط¨ظ†ط§ط، (APP_VERSION + ط±ظ‚ظ… ط§ظ„ط¨ظ†ط§ط،)
const String kAppVersion = String.fromEnvironment('APP_VERSION', defaultValue: 'طھط¬ط±ظٹط¨ظٹ');


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

// ط¨ظٹظ‚ط±ط§ ظ…ظ„ظپ ط§ظ„ط¥ط¹ط¯ط§ط¯ط§طھ (ط§ظ„ظ…ط§ظƒظٹظ†ط§طھ / ظƒظ„ظ…ط§طھ ط§ظ„ظˆط­ط¯ط§طھ / ط£ظƒظˆط§ط¯ ط§ظ„طھط§ط¬ط± / طھطµط­ظٹط­ ط§ظ„ط£ط±ظ‚ط§ظ…). ظ„ظˆ ظ…ظپظٹط´ ظ…ظ„ظپ ط§ظ„ط¥ط¹ط¯ط§ط¯ط§طھ ظپط§ط¶ظٹط©.
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
      if (name.isEmpty) name = 'ظˆط­ط¯ط© $xlRow';
      newTerms[tid] = TerminalInfo(order, xlRow, name);
    }

    final newHints = <MapEntry<List<String>, int>>[];
    for (final r in rowsOf('Name_Hints')) {
      if (r.length < 2) continue;
      final row = _cellInt(r[0]);
      final kws = _cellText(r[1])
          .split(RegExp(r'[,;طŒ|]'))
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
    rebuildTerminalIndexes();
  } catch (e) {
    _dbg('Excel load error: $e');
  }
}

// ط¨ظٹظƒطھط¨ ط§ظ„ظ…ط§ظƒظٹظ†ط§طھ ظˆط§ظ„ظˆط­ط¯ط§طھ ط§ظ„ظ„ظٹ ط§طھط¹ط±ظپطھ (ظ…ظ† ط§ظ„ط¥ظٹطµط§ظ„ط§طھ ط£ظˆ ظ…ظ† ط§ظ„ظ…ظ„ظپ) ظپظٹ ظ…ظ„ظپ ط§ظ„ط¥ط¹ط¯ط§ط¯ط§طھ
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
    final bytes = excel.save();
    if (bytes == null) return false;
    f.writeAsBytesSync(bytes);
    return true;
  } catch (e) {
    _dbg('settings save error: $e');
    return false;
  }
}

// ظ‚ط§ظ„ط¨ ط´ظٹطھ ط§ظ„ظ†طھظٹط¬ط© ط§ظ„ظپط§ط¶ظٹ (ط¹ظ†ط§ظˆظٹظ† ط¨ط³) ظ„ظˆ ظ…ظپظٹط´ ظ‚ط§ظ„ط¨ ظ…ط¶ط§ظپ ظ„ظ„طھط·ط¨ظٹظ‚ ظˆظ„ط§ ظپظٹ ط§ظ„ظپظˆظ„ط¯ط±
List<int>? _blankTemplateBytes() {
  try {
    final e = ex.Excel.createExcel();
    final name = e.getDefaultSheet() ?? 'Sheet1';
    e.appendRow(name, <ex.CellValue?>[
      null,
      ex.TextCellValue('ط§ظ„ظˆط­ط¯ط©'),
      ex.TextCellValue('ط¹ط¯ط¯ ط§ظ„ظ…ط§ظƒظٹظ†ط§طھ'),
      ex.TextCellValue('ط§ظ„ظ…ط¨ظ„ط؛'),
      ex.TextCellValue('ظƒط§ط´'),
      ex.TextCellValue('ط§ظ„ط¥ط¬ظ…ط§ظ„ظٹ'),
    ]);
    return e.save();
  } catch (_) {
    return null;
  }
}

// ط¨ظٹطھظ†ط¯ظ‡ ط£ظˆظ„ ظ…ط§ ط§ظ„طھط·ط¨ظٹظ‚ ظٹظپطھط­ (ظˆظ„ظ…ط§ ظٹط±ط¬ط¹ ظ…ظ† ط´ط§ط´ط© ط§ظ„ط¥ط°ظ†): ط¨ظٹظ†ط²ظ‘ظ„ ط´ظٹطھ ط§ظ„ط¥ط¹ط¯ط§ط¯ط§طھ ظˆط´ظٹطھ ط§ظ„ظ†طھظٹط¬ط© طھظ„ظ‚ط§ط¦ظٹ ظˆظٹط­ظ…ظ‘ظ„ ط§ظ„ط¥ط¹ط¯ط§ط¯ط§طھ
Future<void> prepareDefaultFiles() async {
  try {
    if (await Permission.manageExternalStorage.isGranted) {
      await ensureDefaultFiles();
      await ensureSettingsFile();
    }
  } catch (_) {}
  await loadSettingsFromExcel();
}

// ط£ظˆظ„ ظ…ط±ط©: ط¨ظٹظ†ط³ط® ظ…ظ„ظپظٹ ط§ظ„ط¥ط¹ط¯ط§ط¯ط§طھ ظˆط§ظ„ظ‚ط§ظ„ط¨ (ط§ظ„ظ…ط±ظپظˆط¹ظٹظ† ظ…ط¹ ط§ظ„ظ…ط´ط±ظˆط¹) ظ„ظ…ط¬ظ„ط¯ ط§ظ„طھط·ط¨ظٹظ‚ ظ„ظˆ ظ…ط´ ظ…ظˆط¬ظˆط¯ظٹظ†
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

// ط¨ظٹط¹ظ…ظ„ ظ…ظ„ظپ ط§ظ„ط¥ط¹ط¯ط§ط¯ط§طھ ظپط§ط¶ظٹ (ط¨ط§ظ„ط¹ظ†ط§ظˆظٹظ† ط¨ط³) ظ„ظˆ ظ…ط´ ظ…ظˆط¬ظˆط¯طŒ ط¹ط´ط§ظ† ط§ظ„ظ…ط³طھط®ط¯ظ… ظٹظ…ظ„ط§ظ‡
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
      'Terminals: Order = ط±ظ‚ظ… ط§ظ„ظ…ط§ظƒظٹظ†ط© (1,2,3...) | TID = ط±ظ‚ظ… ط§ظ„طھط±ظ…ظ†ظ„ ط§ظ„ظ…ظƒطھظˆط¨ ط¹ظ„ظ‰ ط§ظ„ط¥ظٹطµط§ظ„ | Row = ط±ظ‚ظ… طµظپ ط§ظ„ظˆط­ط¯ط© ظپظٹ ظ…ظ„ظپ ط§ظ„ط¥ظƒط³ظٹظ„ | Unit_Name = ط§ط³ظ… ط§ظ„ظˆط­ط¯ط©',
      'Name_Hints: ظƒظ„ظ…ط§طھ (ط¨ط§ظ„ط¥ظ†ط¬ظ„ظٹط²ظٹطŒ ظ…ظپطµظˆظ„ط© ط¨ظپط§طµظ„ط©) ظ„ظˆ ط¸ظ‡ط±طھ ظپظٹ ط§ظ„ط¥ظٹطµط§ظ„ طھطھط¹ط±ظپ ط§ظ„ظˆط­ط¯ط©',
      'Merchant_Codes: ظƒظˆط¯ ط§ظ„طھط§ط¬ط± -> ط±ظ‚ظ… ط§ظ„طµظپ',
      'TID_Fixes: ط±ظ‚ظ… طھط±ظ…ظ†ظ„ ط¨طھظ‚ط±ط£ظ‡ OCR ط؛ظ„ط· -> ط§ظ„ط±ظ‚ظ… ط§ظ„طµط­',
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
            label: tr('ط§ظ„طھظ‚ظپظٹظ„'),
          ),
          NavigationDestination(
            icon: _badged(machBadge, Icons.folder_copy),
            label: tr('ط¥ظٹطµط§ظ„ط§طھ ط§ظ„ط¯ظپط¹'),
          ),
        ],
      ),
    );
  }
}

const String kAppName = 'A2R POS Public';
const String kAppAuthor = 'Ahmed Ramadan - Alexandria';

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ ط§ظ„ط¥ط¹ط¯ط§ط¯ط§طھ: ط§ظ„ظ„ط؛ط© + ط§ظ„ط¯ط§ط±ظƒ ظ…ظˆط¯ â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
class AppSettings extends ChangeNotifier with WidgetsBindingObserver {
  ThemeMode themeMode = ThemeMode.system; // system = ط­ط³ط¨ ط§ظ„ظ†ط¸ط§ظ…
  String? langCode; // null = ط­ط³ط¨ ظ„ط؛ط© ط§ظ„ظ†ط¸ط§ظ…
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
    if (langCode == null) notifyListeners(); // ط¨ظ†طھط¨ط¹ ظ„ط؛ط© ط§ظ„ظ†ط¸ط§ظ…
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

  // ط£ظٹ ظˆظٹط¯ط¬طھ ط¨طھظ†ط¯ظ‡ ط¹ظ„ظٹظ‡ط§ ط¨طھطھط³ط¬ظ‘ظ„ ظˆط¨طھطھط­ط¯ظ‘ط« ظ„ظˆط­ط¯ظ‡ط§ ظ„ظ…ط§ ط§ظ„ظ„ط؛ط© طھطھط؛ظٹط±
  static AppSettings of(BuildContext context) {
    final s = context.dependOnInheritedWidgetOfExactType<_LangScope>();
    return s?.notifier ?? appSettings;
  }
}

final AppSettings appSettings = AppSettings();

// ظ„ظˆظ† ظ‡ظˆظٹط© ط§ظ„طھط·ط¨ظٹظ‚ (A2R): ط£ط²ط±ظ‚ ظ…ظ„ظƒظٹ. ط¨ظٹطھط¶ط§ظپ ظپظٹ ط¢ط®ط± ط§ظ„ظ‚ط§ظٹظ…ط© ط¹ظ„ط´ط§ظ† ط§ظ„ظ„ظٹ ط§ط®طھط§ط± ظ„ظˆظ† ظ‚ط¨ظ„ ظƒط¯ظ‡ ظٹظپط¶ظ„ ظ„ظˆظ†ظ‡.
const Color kBrandSeed = Color(0xFF2B59C3);
const int kDefaultSeed = 6; // ظ…ظƒط§ظ† kBrandSeed ظپظٹ kSeeds
const List<Color> kSeeds = [Colors.teal, Colors.blue, Colors.indigo, Colors.green, Colors.orange, Colors.purple, kBrandSeed];
const List<String> kSeedNamesAr = ['طھظٹظ„', 'ط£ط²ط±ظ‚', 'ظ†ظٹظ„ظٹ', 'ط£ط®ط¶ط±', 'ط¨ط±طھظ‚ط§ظ„ظٹ', 'ط¨ظ†ظپط³ط¬ظٹ', 'A2R (ط§ظ„ظ‡ظˆظٹط©)'];
const List<String> kSeedNamesEn = ['Teal', 'Blue', 'Indigo', 'Green', 'Orange', 'Purple', 'A2R (brand)'];

// â”€â”€â”€â”€â”€â”€â”€â”€â”€ ط­ط§ظ„ط© ظ…ط´طھط±ظƒط© ط¨ظٹظ† ط§ظ„طھط§ط¨ظٹظ† â”€â”€â”€â”€â”€â”€â”€â”€â”€
final ValueNotifier<Map<int, int>> closingTxn = ValueNotifier<Map<int, int>>(<int, int>{}); // ط±ظ‚ظ… ط§ظ„ظ…ط§ظƒظٹظ†ط© -> ط¹ط¯ط¯ ظ…ط¹ط§ظ…ظ„ط§طھ ط§ظ„ط¨ط§طھط´
final ValueNotifier<int> wipeSignal = ValueNotifier<int>(0); // ط¨ظٹطھط؛ظٹط± ظ„ظ…ط§ "ظ…ط³ط­ ظƒظ„ ط­ط§ط¬ط©" ظٹط®ظ„طµ
bool homeBusyG = false;
bool machBusyG = false;

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

// طھط£ظƒظٹط¯ ظ‚ط¨ظ„ ط§ظ„ط¥ط±ط³ط§ظ„ ظ„ظˆ ظپظٹظ‡ ط­ط§ط¬ط§طھ ظ„ط³ظ‡ ظ…ظپطھظˆط­ط©
Future<bool> confirmSend(BuildContext context, {int review = 0, int missing = 0, int pending = 0}) async {
  if (review == 0 && missing == 0 && pending == 0) return true;
  final ar = AppSettings.of(context).isAr;
  final lines = <String>[
    if (review > 0) (ar ? 'â€¢ $review ظ…ط­طھط§ط¬ ظ…ط±ط§ط¬ط¹ط©' : 'â€¢ $review need review'),
    if (missing > 0) (ar ? 'â€¢ $missing ظ…ط§ظƒظٹظ†ط© ظ†ط§ظ‚طµط©' : 'â€¢ $missing missing machine(s)'),
    if (pending > 0) (ar ? 'â€¢ $pending ظ„ط³ظ‡ ظ…ط§ ط§طھط¹ط¯ظ‘ط´' : 'â€¢ $pending not counted yet'),
  ];
  final r = await showDialog<bool>(
    context: context,
    builder: (ctx) => AlertDialog(
      title: w.Text(ar ? 'ظ„ط³ظ‡ ظپظٹظ‡ ط­ط§ط¬ط§طھ ظ…ظپطھظˆط­ط©' : 'Some things are still open'),
      content: w.Text('${lines.join('\n')}\n\n${ar ? 'طھط¨ط¹طھ ط¨ط±ط¶ظ‡طں' : 'Send anyway?'}'),
      actions: [
        TextButton(onPressed: () => Navigator.pop(ctx, false), child: w.Text(ar ? 'ط±ط¬ظˆط¹' : 'Back')),
        FilledButton(onPressed: () => Navigator.pop(ctx, true), child: w.Text(ar ? 'ط§ط¨ط¹طھ ط¨ط±ط¶ظ‡' : 'Send anyway')),
      ],
    ),
  );
  return r == true;
}

// ظ…ط³ط­ ظƒظ„ ط­ط§ط¬ط©: ظ…ظ„ظپط§طھ ط§ظ„ط¥ط®ط±ط§ط¬ + ط§ظ„طµظˆط± + ظ†طھط§ظٹط¬ ط§ظ„ظپط­طµ + ط§ظ„طھطµط­ظٹط­ط§طھ + ط§ظ„ظ„ظˆط¬ + ط§ظ„ظƒط§ط´ ظˆط§ظ„ظ…ظ„ظپط§طھ ط§ظ„ظ…ط¤ظ‚طھط© (ط§ظ„ط¥ط¹ط¯ط§ط¯ط§طھ ظˆط§ظ„ظ„ط؛ط© ظˆط§ظ„ط«ظٹظ… ط¨ظٹظپط¶ظ„ظˆط§)
Future<void> wipeEverything(BuildContext context) async {
  final ar = AppSettings.of(context).isAr;
  final m = ScaffoldMessenger.of(context);
  if (homeBusyG || machBusyG) {
    m.showSnackBar(SnackBar(
        content: w.Text(ar ? 'ط§ط³طھظ†ظ‰ ظ„ط­ط¯ ظ…ط§ ط§ظ„ظ…ط¹ط§ظ„ط¬ط© طھط®ظ„طµ ط£ظˆ ظˆظ‚ظ‘ظپظ‡ط§ ط§ظ„ط£ظˆظ„.' : 'Wait for processing to finish or stop it first.')));
    return;
  }
  final ok = await showDialog<bool>(
    context: context,
    builder: (ctx) => AlertDialog(
      title: w.Text(ar ? 'ظ…ط³ط­ ظƒظ„ ط­ط§ط¬ط©طں' : 'Wipe everything?'),
      content: w.Text(ar
          ? 'ظ‡ظٹطھظ…ط³ط­ ظ†ظ‡ط§ط¦ظٹ:\nâ€¢ ظ…ظ„ظپط§طھ ط§ظ„ط¥ط®ط±ط§ط¬ (ط§ظ„ط¥ظƒط³ظٹظ„ ظˆط§ظ„ظ€ PDF)\nâ€¢ ط§ظ„طµظˆط± ط§ظ„ظ…ط±طھط¨ط© ظˆط§ظ„ظ†ط³ط® ط§ظ„ظ…طھط³ظ…ظٹط© ظ„ط¥ظٹطµط§ظ„ط§طھ ط§ظ„ط¯ظپط¹\nâ€¢ ظ†طھط§ظٹط¬ ط§ظ„ظپط­طµ ظˆط§ظ„طھطµط­ظٹط­ط§طھ ط§ظ„ظ…ط­ظپظˆط¸ط©\nâ€¢ ط§ظ„ظ„ظˆط¬\nâ€¢ ط§ظ„ظƒط§ط´ ظˆط§ظ„ظ…ظ„ظپط§طھ ط§ظ„ظ…ط¤ظ‚طھط©\n\nظ…ط´ ظ‡ظٹطھظ…ط³: ظ…ظ„ظپط§طھ ظˆط§طھط³ط§ط¨ ط§ظ„ط£طµظ„ظٹط©طŒ ظˆط§ظ„ط¥ط¹ط¯ط§ط¯ط§طھ (ط§ظ„ظ„ط؛ط© ظˆط§ظ„ط«ظٹظ… ظˆط§ظ„ظپظˆظ„ط¯ط±ط§طھ).'
          : 'This will permanently delete:\nâ€¢ Output files (Excel and PDF)\nâ€¢ Sorted images and renamed payment receipt copies\nâ€¢ Scan results and saved corrections\nâ€¢ The log\nâ€¢ Cache and temporary files\n\nNot touched: original WhatsApp files and settings (language, theme, folders).'),
      actions: [
        TextButton(onPressed: () => Navigator.pop(ctx, false), child: w.Text(ar ? 'ط¥ظ„ط؛ط§ط،' : 'Cancel')),
        FilledButton(
          style: FilledButton.styleFrom(backgroundColor: Colors.red),
          onPressed: () => Navigator.pop(ctx, true),
          child: w.Text(ar ? 'ط§ظ…ط³ط­ ظƒظ„ ط­ط§ط¬ط©' : 'Wipe everything'),
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
    final l = n.toLowerCase();
    return n == 'review_report.txt' ||
        n == 'ocr_debug.txt' ||
        _isOutputFileName(n);
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
  wipeTree(defaultMDest);
  // ظ†ط³ط® ط¥ظٹطµط§ظ„ط§طھ ط§ظ„ط¯ظپط¹ ط§ظ„ظ…طھط³ظ…ظٹط© ظ„ظˆ ط§ظ„ظپظˆظ„ط¯ط± ظ…ط®طµطµ
  try {
    final dest = loadMOut();
    if (dest != defaultMDest) {
      for (final r in loadMState()) {
        final p = r.openPath;
        if (r.newName != null && p != null && p != r.srcPath && p.startsWith(dest)) {
          final f = File(p);
          if (f.existsSync()) delFile(f);
        }
        r.ctl.dispose();
      }
    }
  } catch (_) {}
  for (final p in [fixPath, mStatePath, '$baseDir/id_debug.txt']) {
    final f = File(p);
    if (f.existsSync()) delFile(f);
  }
  try {
    final cfg = _readCfg();
    cfg.remove('doneM');
    _writeCfg(cfg);
  } catch (_) {}
  // ط§ظ„ظƒط§ط´ ظˆط§ظ„ظ…ظ„ظپط§طھ ط§ظ„ظ…ط¤ظ‚طھط© ط¨طھط§ط¹ط© ط§ظ„طھط·ط¨ظٹظ‚
  try {
    for (final e in Directory.systemTemp.listSync(followLinks: false)) {
      if (e is File) {
        delFile(e);
      } else if (e is Directory) {
        wipeTree(e.path);
      }
    }
  } catch (_) {}
  imageCache.clear();
  imageCache.clearLiveImages();
  closingTxn.value = <int, int>{};
  wipeSignal.value++;
  final mb = (bytes / 1048576).toStringAsFixed(1);
  m.showSnackBar(SnackBar(
      content: w.Text(ar ? 'طھظ… ط§ظ„ظ…ط³ط­: $files ظ…ظ„ظپ ($mb ظ…ظٹط¬ط§) ًں§¹' : 'Wiped: $files files ($mb MB) ًں§¹')));
}

final ValueNotifier<int> sharedDays = ValueNotifier<int>(0); // ط§ظ„ظپطھط±ط© (ط§ظ„ظٹظˆظ…/ظٹظˆظ…ظٹظ†/...) ظˆط§ط­ط¯ط© ظ„ظ„طھط§ط¨ظٹظ†
final ValueNotifier<int> homeBadge = ValueNotifier<int>(0); // ط¹ط¯ط¯ ط§ظ„ظ„ظٹ ظ…ط­طھط§ط¬ ظ…ط±ط§ط¬ط¹ط© ظپظٹ ط§ظ„طھظ‚ظپظٹظ„
final ValueNotifier<int> machBadge = ValueNotifier<int>(0); // ط¹ط¯ط¯ ط§ظ„ظ„ظٹ ظ…ط­طھط§ط¬ ظ…ط±ط§ط¬ط¹ط© ظپظٹ ط¥ظٹطµط§ظ„ط§طھ ط§ظ„ط¯ظپط¹

void _syncBadge(ValueNotifier<int> n, int v) {
  if (n.value == v) return;
  WidgetsBinding.instance.addPostFrameCallback((_) => n.value = v);
}

Future<void> shareDebugLog(BuildContext context) async {
  final m = ScaffoldMessenger.of(context);
  final f = File('$baseDir/id_debug.txt');
  if (!f.existsSync()) {
    m.showSnackBar(const SnackBar(content: Text('ظ…ظپظٹط´ ظ„ظˆط¬ ظ„ط³ظ‡. ط´ط؛ظ‘ظ„ ط§ظ„ظپط­طµ ط§ظ„ط£ظˆظ„.')));
    return;
  }
  await Share.shareXFiles([XFile(f.path, mimeType: 'text/plain')], text: 'ظ„ظˆط¬ ط§ظ„طھط·ط¨ظٹظ‚');
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

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ ط§ظ„طھط±ط¬ظ…ط© (ط¹ط±ط¨ظٹ â†گ ط¥ظ†ط¬ظ„ظٹط²ظٹ) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// ط§ظ„ظ†طµظˆطµ ظ…ظƒطھظˆط¨ط© ط¨ط§ظ„ط¹ط±ط¨ظٹ ظپظٹ ط§ظ„ظƒظˆط¯طŒ ظˆط¨طھطھطھط±ط¬ظ… ظˆظ‚طھ ط§ظ„ط¹ط±ط¶ ظ„ظ…ط§ ط§ظ„ظ„ط؛ط© طھظƒظˆظ† ط¥ظ†ط¬ظ„ظٹط²ظٹ.
// ط£ظٹ ظ†طµ ط¹ط±ط¨ظٹ ظ…ط´ ظ…ظˆط¬ظˆط¯ ظپظٹ ط§ظ„ظ‚ط§ظ…ظˆط³ ط¨ظٹظپط¶ظ„ ط²ظٹ ظ…ط§ ظ‡ظˆ.
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
late final List<RegExp> _ruleRes = [for (final r in _kEnRules) RegExp(r[0])];
late final List<RegExp> _phraseRes = [
  for (final p in _kEnPhrases) RegExp('(?<![$_arL])${RegExp.escape(p[0])}(?![$_arL])')
];

String tr(String s) {
  if (s.isEmpty || appSettings.isAr || !_hasArabic.hasMatch(s)) return s;
  final cached = _trCache[s];
  if (cached != null) return cached;
  final r = _translateEn(s);
  if (_trCache.length > 3000) _trCache.clear();
  _trCache[s] = r;
  return r;
}

String _translateEn(String s) {
  final ex = _kEnExact[s];
  if (ex != null) return ex;
  var out = s;
  for (var i = 0; i < _ruleRes.length; i++) {
    final rep = _kEnRules[i][1];
    out = out.replaceAllMapped(_ruleRes[i], (m) {
      var t = rep;
      for (var g = 1; g <= m.groupCount; g++) {
        t = t.replaceAll('{$g}', m.group(g) ?? '');
      }
      return t;
    });
  }
  for (var i = 0; i < _phraseRes.length; i++) {
    final en = _kEnPhrases[i][1];
    out = out.replaceAllMapped(_phraseRes[i], (_) => en);
  }
  return out.replaceAll('طŒ', ',').replaceAll('طں', '?');
}

const Map<String, String> _kEnExact = {
  'ط§ظ„ظٹظˆظ…': 'Today',
  'ط§ظ„ظƒظ„': 'All',
  'طھظ…ط§ظ…': 'OK',
  'ظ„ط§': 'No',
  'ط¥ظ„ط؛ط§ط،': 'Cancel',
  'ظ…ط³ط­': 'Delete',
  'ط­ظپط¸': 'Save',
  'ط¹ط±ط¶': 'View',
  'ط§ظ„طµط­': 'Correct',
  'طھط؛ظٹظٹط±': 'Change',
  'طھط¬ط±ظٹط¨ظٹ': 'Dev',
  'ظٹظˆظ…ظٹظ†': '2 days',
  'طں': '?',
  'طµط­ âœ“': 'Correct âœ“',
  'ط؛ظ„ط· âœ—': 'Wrong âœ—',
  'ط§ظ„ظ…ط²ظٹط¯': 'More',
  'ط§ظ„ظ…طµط¯ط±': 'Source',
  'ط§ظ„ط¥ط®ط±ط§ط¬': 'Output',
  'ظˆط§طھط³ط§ط¨': 'WhatsApp',
  'ط¥ظƒط³ظٹظ„': 'Excel',
  'طھط¯ظˆظٹط±': 'Rotate',
  'ط§ظ„ط³ط§ط¨ظ‚': 'Previous',
  'ط§ظ„طھط§ظ„ظٹ': 'Next',
  'ط¥ظٹظ‚ط§ظپ': 'Stop',
  'ظ…ط±ط§ط¬ط¹ط©': 'Review',
  'ظ…ط§ظƒظٹظ†ط©': 'Machine',
  'ظ†طھظٹط¬ط©': 'Result'
};

const List<List<String>> _kEnRules = [
  [r'(\d+) ظ…ظ† (\d+)', '{1} of {2}'],
  [r'ط¥ظٹطµط§ظ„ (\d+) ط¨طھط§ط±ظٹط®', 'Receipt {1} dated'],
  [r'ظ†ط§ظ‚طµ (\d+) ظ…ط§ظƒظٹظ†ط©', 'missing {1} machine(s)'],
  [r'ظ…ط±ط§ط¬ط¹ط© (\d+) / (\d+)', 'Review {1} / {2}'],
  [r'ظپط§ط¶ظ„ (\d+) ظ„ط³ظ‡', '{1} remaining']
];

// ظ…ط±طھط¨ط© ظ…ظ† ط§ظ„ط£ط·ظˆظ„ ظ„ظ„ط£ظ‚طµط± ط¹ظ„ط´ط§ظ† ط§ظ„ط¹ط¨ط§ط±ط§طھ ط§ظ„ط·ظˆظٹظ„ط© طھطھط¨ط¯ظ‘ظ„ ط§ظ„ط£ظˆظ„
const List<List<String>> _kEnPhrases = [
  ['Terminal ID ظ…ط´ ظˆط§ط¶ط­طŒ ط§ظ„ظˆط­ط¯ط© ط§طھط­ط¯ط¯طھ ظ…ظ† Merchant ID/ط§ظ„ط§ط³ظ… ط¬ظˆظ‡ ط§ظ„ظ€ PDFطŒ ط±ط§ط¬ط¹', 'Terminal ID unclear; unit determined from Merchant ID/name inside the PDF, review'],
  ['ط§ظ„ظ…ظ„ظپ ط¯ظ‡ ظ…ط´ ظˆط§ط¶ط­ ط¥ظ†ظ‡ ط¥ظٹطµط§ظ„ ظ…ط§ظƒظٹظ†ط© (ظ‚ط±ط§ط،ط© ط¶ط¹ظٹظپط© ظ„ظ„ظ€ Terminal ID)طŒ ط±ط§ط¬ط¹ظ‡', 'This file does not look like a machine receipt (weak Terminal ID reading), review it'],
  ['ط§ظ„ظˆط­ط¯ط© ط§طھط­ط¯ط¯طھ ط¨ط§ظ„ط§ط³ظ…/ط§ظ„ظ…ظٹط±ط´ط§ظ†طھ ظپظ‚ط· - ط§طھط£ظƒط¯ ظ…ظ† ط§ظ„ظ…ط§ظƒظٹظ†ط© ظˆط§ط¶ط؛ط· ط­ظپط¸', 'Unit was determined by name/merchant only - verify the machine and tap Save'],
  ['ط§ط¶ط؛ط· "ط¨ط¯ط، ط§ظ„ظ…ط¹ط§ظ„ط¬ط© ط§ظ„ط´ط§ظ…ظ„ط©" ظ„ط³ط­ط¨ ط§ظ„طµظˆط± ظˆطھظˆظ„ظٹط¯ ط§ظ„ط¥ظƒط³ظٹظ„ ظˆط§ظ„ظ€ PDF.', 'Tap "Run full processing" to collect the images and generate the Excel and PDF.'],
  ['ط§ظ„طھط·ط¨ظٹظ‚ ط§طھط¬ط§ظ‡ظ„ ط§ظ„طµظˆط±ط© ط¯ظٹ ظ„ط£ظ† ط´ظƒظ„ظ‡ط§ ط¥ظٹطµط§ظ„ ط¯ظپط¹ ط£ظˆ ظ…ط³طھظ†ط¯ ظ…ط´ ط¨ط§طھط´', 'The app ignored this image because it looks like a payment receipt or a document, not a batch'],
  ['ط§ط¶ط؛ط· "ظپط­طµ ظˆط¥ط¹ط§ط¯ط© طھط³ظ…ظٹط©" ظ„ظپط­طµ ط¥ظٹطµط§ظ„ط§طھ ط§ظ„ط¯ظپط¹ ظ…ظ† ظ…ط³طھظ†ط¯ط§طھ ظˆط§طھط³ط§ط¨.', 'Tap "Scan & rename" to scan payment receipts from WhatsApp documents.'],
  ['ظ…ظپظٹط´ طµظˆط± (ظپظٹ ط§ظ„ظپطھط±ط© ط§ظ„ظ…ط­ط¯ط¯ط©) ظˆظ„ط§ ظ…ظ„ظپط§طھ PDF ظپظٹ ط§ظ„ظ…ط¬ظ„ط¯ ط§ظ„ظ…ط®طھط§ط±', 'No images (in the selected period) and no PDF files in the selected folder'],
  ['ط§ظ„ط¹ط¯ ظˆظ‚ظپ ط£ظƒطھط± ظ…ظ† ظ…ط±ط© (ط§ظ„طھط·ط¨ظٹظ‚ ط§طھظ‚ظپظ„). ظ…ظ† â‹® ط§ط®طھط§ط± ط¥ط¹ط§ط¯ط© ط§ظ„ط¹ط¯', 'Counting stopped more than once (the app closed). From â‹® choose Recount'],
  ['ظ…ط´ ظ…ط­ط³ظˆط¨ ظپظٹ ط§ظ„ط¥ط¬ظ…ط§ظ„ظٹ ظ„ط­ط¯ ظ…ط§ طھظ‚ط±ط±. ط²ظˆظ‘ظ… ط¹ظ„ظ‰ ط§ظ„ط¨ط§طھط´ ظˆظ‚ط§ط±ظ†', 'Not counted in the total until you decide. Zoom in on the batch and compare'],
  ['ط§طھط¹ظ…ظ„ ظ„ظ‡ط§ PDF ظ…ظ†ظپطµظ„ ظ„ظƒظ„ ظٹظˆظ…طŒ ظˆظ…ط´ ط¯ط§ط®ظ„ط© ظپظٹ ط¥ط¬ظ…ط§ظ„ظٹط§طھ ظٹظˆظ…', 'a separate PDF was created for each day, and they are not included in the totals of day'],
  ['ط²ظˆظ‘ظ… ظˆط´ظˆظپ: ظ„ظˆ ط¨ط§طھط´ ط­ظ‚ظٹظ‚ظٹ ط§ط¯ط®ظ„ظ‡ ظˆظƒظ…ظ‘ظ„ ط§ظ„ظ…ط¨ظ„ط؛ ظˆط§ظ„ظ…ط§ظƒظٹظ†ط©', 'Zoom in and check: if it is a real batch, enter it and fill in the amount and machine'],
  ['ط§ظ„ظ€ PDF ط§طھط¹ط±ظ‘ظپ ط¹ظ„ظ‰ ظ†ظپط³ ظ…ط§ظƒظٹظ†ط© ظ…ظ„ظپ طھط§ظ†ظٹطŒ ط§طھط³ظ…ظ‘ظ‰ ظ…ط§ظƒظٹظ†ط©', 'The PDF was identified as the same machine as another file, named machine'],
  ['ظˆظ‡طھطھظ…ط³ط­ ظ†طھظٹط¬ط© ط§ظ„ظپط­طµ. (ظ…ظ„ظپط§طھ ظˆط§طھط³ط§ط¨ ط§ظ„ط£طµظ„ظٹط© ظ…ط´ ظ‡طھطھظ…ط³)', 'The scan result will also be deleted. (Original WhatsApp files will not be touched)'],
  ['ط¸ظ‡ط±طھ ظپظٹ طµظˆط±ط© طھط§ظ†ظٹط© ط¨ظ…ط¨ظ„ط؛ ظ…ط®طھظ„ظپ - ط§طھط­ط³ط¨طھ ظƒط¨ط§طھط´ ط¥ط¶ط§ظپظٹ', 'appeared in another image with a different amount - counted as an extra batch'],
  ['ظپظٹظ‡ ط¥ظٹطµط§ظ„ ط¥ظ‚ظپط§ظ„/ظ…ظ„ط®طµ/ط¥ط¹ط§ط¯ط© ط·ط¨ط§ط¹ط© (ط§طھط¬ط§ظ‡ظ„ ظˆظ…ط´ ظ…ط¹ط¯ظˆط¯)', 'Contains a closing/summary/reprint receipt (ignored, not counted)'],
  ['ظپظˆظ„ط¯ط± ط§ظ„ط¥ط®ط±ط§ط¬ ط¯ظ‡ ط¹ط§ظ…طŒ ظپظ‡ظٹطھظ…ط³ط­ ط¨ط³ ظ…ظ„ظپط§طھ ط§ظ„طھط·ط¨ظٹظ‚ ظ…ظ†ظ‡', 'this output folder is shared, so only the app files will be deleted from it'],
  ['ط¹ط¯ط¯ ط¥ظٹطµط§ظ„ط§طھ ط§ظ„ط¯ظپط¹: ظ„ط³ظ‡ ظ…ط§ ط§طھط¹ط¯ظ‘ط´ (ط§ط¶ط؛ط· ط¨ط¯ط، ط§ظ„ط¹ط¯)', 'Payment receipts count: not counted yet (tap Start counting)'],
  ['ظ…ط´ ظ„ط§ظ‚ظٹ ظپظˆظ„ط¯ط± ظ…ط³طھظ†ط¯ط§طھ ظˆط§طھط³ط§ط¨. ط§ط®طھط§ط± ظپظˆظ„ط¯ط± ط¨ط¥ظٹط¯ظƒ', 'Could not find the WhatsApp documents folder. Pick a folder manually'],
  ['ط§طھط¬ط§ظ‡ظ„طھظ‡: ط´ظƒظ„ظ‡ ط¥ظٹطµط§ظ„ ط¯ظپط¹/ظ…ط³طھظ†ط¯ ظ…ط´ ط¨ط§طھط´ - ط±ط§ط¬ط¹ظ‡', 'Ignored: looks like a payment receipt/document, not a batch - review it'],
  ['ظ…ظ„ظپ ظ„ط³ظ‡ ظ…ط§ ط§طھط¹ط¯ظ‘ط´طŒ ط§ط¶ط؛ط· "ظپط­طµ" ظˆظ‡ظٹظƒظ…ظ‘ظ„ ظ…ظ† ظ…ظƒط§ظ†ظ‡', 'file(s) not counted yet, tap "Scan" and it will continue where it left off'],
  ['ظٹط¬ط¨ ظ…ظ†ط­ ط¥ط°ظ† ط§ظ„ظˆطµظˆظ„ ط¥ظ„ظ‰ ط¬ظ…ظٹط¹ ط§ظ„ظ…ظ„ظپط§طھ ظ„ظ„ظ…طھط§ط¨ط¹ط©', 'You must grant all-files access to continue'],
  ['ظ„ط§ طھظˆط¬ط¯ طµظˆط± ظپظٹ ظ…ط¬ظ„ط¯ ط§ظ„ظˆط§طھط³ط§ط¨ ط¨ط§ظ„ظپطھط±ط© ط§ظ„ظ…ط­ط¯ط¯ط©', 'No images in the WhatsApp folder for the selected period'],
  ['طھط¹ط°ط± ط§ظ„ط¹ط«ظˆط± ط¹ظ„ظ‰ ظ…ط¬ظ„ط¯ طµظˆط± ط§ظ„ظˆط§طھط³ط§ط¨ طھظ„ظ‚ط§ط¦ظٹط§ظ‹', 'Could not find the WhatsApp images folder automatically'],
  ['ظ…ط´ ظ‚ط§ط¯ط± ط£ظ‚ط±ط§ ط§ظ„ظ…ط¨ظ„ط؛ ط£ظˆ ط§ظ„ظˆط­ط¯ط© - ط§ط¶ط؛ط· طھط¹ط¯ظٹظ„', 'Cannot read the amount or the unit - tap Edit'],
  ['ظ…ظ…ظƒظ† ط¥ظٹطµط§ظ„ط§طھ ظ†ط§ظ‚طµط© ط£ظˆ ظ…ط´ ط§طھظ‚ط±طھطŒ ط±ط§ط¬ط¹ ط§ظ„ظ…ظ„ظپ', 'possibly missing receipts or unread ones, check the file'],
  ['ظ…ظ„ظپ ظ„ط³ظ‡ ظ…ط§ ط§طھط¹ط¯ظ‘ط´. ط§ط¶ط؛ط· "ط¨ط¯ط، ط§ظ„ط¹ط¯" ظ„ظ…ط§ طھط­ط¨', 'file(s) not counted yet. Tap "Start counting" when you want'],
  ['ط§ظ„ظ…ط§ظƒظٹظ†ط©/ط§ظ„ظˆط­ط¯ط© ط§طھط­ط¯ط¯طھ طھظ„ظ‚ط§ط¦ظٹط§ظ‹ - ط§طھط£ظƒط¯', 'Machine/unit was determined automatically - verify'],
  ['ظ…ط´ظƒظˆظƒ ط¥ظ†ظ‡ ظ…ظƒط±ط± - ظ…ط´ ظ…ط­ط³ظˆط¨ ظ„ط­ط¯ ظ…ط§ طھط±ط§ط¬ط¹ظ‡', 'Possible duplicate - not counted until reviewed'],
  ['طµظˆط±ط©طŒ ظˆظ„ظ… ظٹطھظ… ط§ظ„ط¹ط«ظˆط± ط¹ظ„ظ‰ ط¥ظٹطµط§ظ„ط§طھ ظ…ط·ط§ط¨ظ‚ط©', 'image(s), and no matching receipts were found'],
  ['طµظپط­ط© ظپط´ظ„طھ ظ‚ط±ط§ط،طھظ‡ط§طŒ ط§ظ„ط¹ط¯ط¯ ظ…ظ…ظƒظ† ظٹظƒظˆظ† ظ†ط§ظ‚طµ', 'page(s) failed to read, the count may be incomplete'],
  ['ظƒظ„ ط§ظ„ظ…ظ„ظپط§طھ ظپظٹ ط§ظ„ظپطھط±ط© ط¯ظٹ ط§طھظپط­طµطھ ظ‚ط¨ظ„ ظƒط¯ظ‡', 'All files in this period were already scanned'],
  ['ط¬ط§ط±ظٹ طھظˆظ„ظٹط¯ ط§ظ„ظ…ظ„ظپط§طھ ظˆطھط­ط¯ظٹط« ط§ظ„ط´ظٹطھط§طھ...', 'Generating files and updating sheets...'],
  ['ظ„ط§ط²ظ… طھط³ظ…ط­ ط¨ط§ظ„ظˆطµظˆظ„ ظ„ظƒظ„ ط§ظ„ظ…ظ„ظپط§طھ ط§ظ„ط£ظˆظ„', 'You need to allow access to all files first'],
  ['ط§ظ„ظ…ظ„ظپ ظ…ط´ ظ…ظˆط¬ظˆط¯. ط´ط؛ظ‘ظ„ ط§ظ„ظ…ط¹ط§ظ„ط¬ط© ط§ظ„ط£ظˆظ„', 'File not found. Run the processing first'],
  ['ظ…ط±ط§ط¬ط¹ط© ط§ظ„طھظƒط±ط§ط± ظˆطھط±ظ‚ظٹظ… ط§ظ„ظ…ط§ظƒظٹظ†ط§طھ...', 'Checking duplicates and numbering machines...'],
  ['ط§ظ„ظˆط­ط¯ط© ط§طھط­ط¯ط¯طھ ط¨ط§ظ„ط§ط³ظ…/ط§ظ„ظ…ظٹط±ط´ط§ظ†طھ ظپظ‚ط·', 'Unit was determined by name/merchant only'],
  ['ط§ط³ط­ط¨ ط£ظٹ ظ…ظ„ظپ ظ„طھط¹ظ„ظٹظ…ظ‡: طµط­ âœ“ ط£ظˆ ط؛ظ„ط· âœ—', 'Swipe any file to mark it: correct âœ“ or wrong âœ—'],
  ['ظ…ظ‚ط¯ط±طھط´ ط£ظپطھط­ ط§ظ„ظپظˆظ„ط¯ط±. ط§ظ„ظ…ط³ط§ط± ط§طھظ†ط³ط®', 'Could not open the folder. Path copied'],
  ['طھظ‚ط±ظٹط± ط§ظ„ط¥ظٹطµط§ظ„ط§طھ (ط§ظ„ط±ط³ظ…ظٹ)', 'Receipts report (official)'],
  ['ظ„ظ… ظٹطھظ… ط¥ظ†ط´ط§ط، ظ…ظ„ظپط§طھ ط¨ط¹ط¯ ظ„ظ…ط´ط§ط±ظƒطھظ‡ط§', 'No files have been created yet to share'],
  ['ظˆظ„ط§ ظ…ط¶ط§ظپ ظ„ظ„طھط·ط¨ظٹظ‚)طŒ ظپظ…طھط¹ظ…ظ„ط´ ط¥ظƒط³ظٹظ„', 'nor added to the app), so no Excel was created'],
  ['ظ…ظپظٹط´ طµظˆط±. ط¨ط¯ظˆظ‘ط± ط¹ظ„ظ‰ ظ…ظ„ظپ PDF ط§ط³ظ…ظ‡', 'No images. Looking for a PDF file named'],
  ['ظ…ط´ ظ‚ط§ط¯ط± ط£ط¹ط±ظپ ط§ظ„ظ…ط§ظƒظٹظ†ط© ظ…ظ† ط§ظ„ظ€ PDF', 'Cannot identify the machine from the PDF'],
  ['ظ…ظ„ظپط§طھ ط§ظ„ظ…ط±ط§ط¬ط¹ط© ط؛ظٹط± ظ…طھط§ط­ط© ط­ط§ظ„ظٹط§ظ‹', 'Review files are not available right now'],
  ['ظ…ط´ظƒظˆظƒ ط¥ظ†ظ‡ ظ…ظƒط±ط± (ظ†ظپط³ ظ…ط¨ظ„ط؛ ظ…ط§ظƒظٹظ†ط©', 'Possible duplicate (same amount as machine'],
  ['ظ…ط´ ظ‚ط§ط¯ط± ط£ط¹ط±ظپ ط§ظ„طھط§ط±ظٹط® ظ…ظ† ط§ظ„ظ€ PDF', 'Cannot identify the date from the PDF'],
  ['ط§ط³طھط¨ط¹ط¯طھظ‡ ظ‚ط¨ظ„ ظƒط¯ظ‡ - طھظ‚ط¯ط± طھط±ط¬ظ‘ط¹ظ‡', 'excluded earlier - you can restore it'],
  ['ط§ظ„ظ…ط±ط´ط­ ط§ظ„طھط§ظ†ظٹ ظ…ظ† ظ‚ط±ط§ط،ط© ط§ظ„ظ€ PDF', 'second candidate from reading the PDF'],
  ['ظ…ظپظٹط´ ظ„ظˆط¬ ظ„ط³ظ‡. ط´ط؛ظ‘ظ„ ط§ظ„ظپط­طµ ط§ظ„ط£ظˆظ„', 'No log yet. Run the scan first'],
  ['ط¬ط§ط±ظٹ ط§ظ„ط¨ط­ط« ط¹ظ† ظ…ط¬ظ„ط¯ط§طھ ط§ظ„طµظˆط±...', 'Searching for image folders...'],
  ['ط¥ط¹ط§ط¯ط© طھط³ظ…ظٹط© ط§ظ„طµظˆط± ظˆطھط±طھظٹط¨ظ‡ط§...', 'Renaming and sorting images...'],
  ['ط§ط¶ط؛ط· "ط¨ط¯ط، ط§ظ„ط¹ط¯" ظ„ط¹ط¯ظ‘ ط¥ظٹطµط§ظ„ط§طھظ‡', 'Tap "Start counting" to count its receipts'],
  ['طµظپط­ط§طھ ط§ظ„ط·ط±ظٹظ‚طھظٹظ† ط§ط®طھظ„ظپظˆط§ ظپظٹظ‡ط§:', 'Pages where the two methods differ:'],
  ['ط§طھط¹ظ„ظ‘ظ… ظ„ظ„ط¹ط¯. ط§ط¶ط؛ط· "ط¨ط¯ط، ط§ظ„ط¹ط¯".', 'Marked for counting. Tap "Start counting".'],
  ['ظ…ط´ظƒظˆظƒ ط¥ظ†ظ‡ ظ…ظƒط±ط±: ظ†ظپط³ ط§ظ„ظ…ط§ظƒظٹظ†ط©', 'Possible duplicate: same machine'],
  ['ط§ط¶ط؛ط· "ط¨ط¯ط، ط§ظ„ط¹ط¯" ظ„ط¹ط¯ظ‘ ط¥ظٹطµط§ظ„ط§طھ', 'Tap "Start counting" to count receipts of'],
  ['طµظپط­ط§طھ ظ…ط§ ظ„ظ‚ظٹطھط´ ظپظٹظ‡ط§ ط¥ظٹطµط§ظ„ط§طھ:', 'Pages with no receipts found:'],
  ['ط§ظ„ط¹ط¯: ط³ط±ظٹط¹ â†گ ط§ظ„ط¯ظ‚ظٹظ‚ (ط§ظ„ظ‚ط¯ظٹظ…)', 'Counting: fast, then precise (old)'],
  ['ط¹ط¯ط¯ ط¥ظٹطµط§ظ„ط§طھ ط§ظ„ط¯ظپط¹: ط¨ظٹطھط¹ط¯ظ‘...', 'Payment receipts count: counting...'],
  ['طھظ… طھظˆظ„ظٹط¯ ظ…ظ„ظپ ط§ظ„ظ€ PDF ط§ظ„ظ…ط¬ظ…ط¹', 'Combined PDF file generated'],
  ['ط±ظ‚ظ… ط§ظ„ظ…ط§ظƒظٹظ†ط© ط§طھط­ط¯ط¯ طھظ„ظ‚ط§ط¦ظٹط§ظ‹', 'Machine number was determined automatically'],
  ['ظ…ط´ ظ…ظƒط±ط± - ط§ط­طھط³ط¨ظ‡ ط¨ط§طھط´ ط¥ط¶ط§ظپظٹ', 'Not a duplicate - count it as an extra batch'],
  ['ط¥ظٹطµط§ظ„ (ط±ظ‚ظ…ظ‡ ظ…ط´ ظˆط§ط¶ط­) ط¨طھط§ط±ظٹط®', 'Receipt (number unclear) dated'],
  ['ظ…ظپظٹط´ ظ…ظ„ظپط§طھ ظ…ط³ظ…ظ‘ط§ط© ظ„ظ…ط´ط§ط±ظƒطھظ‡ط§', 'No named files to share'],
  ['ظ…ظپظٹط´ ظ…ظ„ظپط§طھ PDF ظپظٹ ط§ظ„ظپطھط±ط© ط¯ظٹ', 'No PDF files in this period'],
  ['ظ…ط´ ظ‚ط§ط¯ط± ط£ظƒطھط¨ ظپظٹ ط§ظ„ظپظˆظ„ط¯ط± ط¯ظ‡', 'Cannot write to this folder'],
  ['ظ…ظپظٹط´ ط¥ظٹطµط§ظ„ط§طھ ظ…ط­طھط§ط¬ط© ظ…ط±ط§ط¬ط¹ط©', 'No receipts need review'],
  ['طµظˆط±ط© ظ…طھط¬ط§ظ‡ظ„ط© - ظ…ط­طھط§ط¬ط© ظ‚ط±ط§ط±', 'Ignored image - needs a decision'],
  ['ظ…ظپظٹط´ ط¨ط±ظ†ط§ظ…ط¬ ظٹظپطھط­ ط§ظ„ظ…ظ„ظپ ط¯ظ‡', 'No app can open this file'],
  ['ظپظٹظ‡ ط¥ظٹطµط§ظ„ط§طھ ط¨طھظˆط§ط±ظٹط® طھط§ظ†ظٹط©', 'There are receipts with other dates'],
  ['ظپط­طµ ط§ظ„طµظˆط± ط§ظ„ظ…ظ‚ظ„ظˆط¨ط© (ط£ط¨ط·ط£)', 'Check rotated images (slower)'],
  ['ط¬ط§ط±ظٹ ظپط­طµ ط¥ظٹطµط§ظ„ط§طھ ط§ظ„ط¯ظپط¹...', 'Scanning payment receipts...'],
  ['ظˆط§طھط³ط§ط¨ - ظ…ط³طھظ†ط¯ط§طھ (طھظ„ظ‚ط§ط¦ظٹ)', 'WhatsApp - documents (auto)'],
  ['طھظ‚ط±ظٹط± ط§ظ„ط¥ظٹطµط§ظ„ط§طھ', 'Receipts report'],
  ['ط¬ط§ط±ظٹ طھط¬ظ‡ظٹط² طµظپط­ط§طھ ط§ظ„ظ€ PDF', 'Preparing PDF pages'],
  ['ط§ظ„ط¥ظٹطµط§ظ„ ط¯ظ‡ ظ…ط³طھط¨ط¹ط¯ (ظ…ظƒط±ط±)', 'This receipt is excluded (duplicate)'],
  ['ظ…ظپظٹط´ ط¨ط±ظ†ط§ظ…ط¬ ظٹظپطھط­ ط§ظ„ظ€ PDF', 'No app can open the PDF'],
  ['ظ…ط´ ظ‚ط§ط¯ط± ط£ظپطھط­ ط§ظ„ظپظˆظ„ط¯ط± ط¯ظ‡', 'Cannot open this folder'],
  ['ط§ظ†طھظ‡ظ‰ ظپط­طµ ط¥ظٹطµط§ظ„ط§طھ ط§ظ„ط¯ظپط¹', 'Payment receipts scan finished'],
  ['ط¬ط§ط±ظٹ ط§ظ„ط¨ط­ط« ط¹ظ† ظ…ظ„ظپط§طھ PDF', 'Searching for PDF files'],
  ['ظ…ط´ ظ…ط­ط³ظˆط¨ ظ„ط­ط¯ ظ…ط§ طھط±ط§ط¬ط¹ظ‡', 'not counted until you review it'],
  ['ط¥ط¬ظ…ط§ظ„ظٹ ط¨ط¹ظٹط¯ ط¹ظ† ط§ظ„ظ…ط¹طھط§ط¯', 'Total far from usual'],
  ['ط§ظ„ظ‚ط±ط§ط،ط© ط§ظ„طھط§ظ†ظٹط© ظ„ظ„ظ…ط¨ظ„ط؛', 'The second amount reading'],
  ['طµظˆط±ط© ظ…ظ† ط§ظ„طµظˆط± ط§ظ„ظ…ط±طھط¨ط©', 'image(s) from the sorted images'],
  ['ط¬ط§ط±ظٹ ظپط­طµ ط§ظ„ط¥ظٹطµط§ظ„ط§طھ...', 'Scanning receipts...'],
  ['ظˆظ„ط§ ظٹظˆط¬ط¯ ظ…ظ„ظپ PDF ط¨ط§ط³ظ…', 'and there is no PDF file named'],
  ['ظ„ظˆ ط§ظ„ظ…ط¨ظ„ط؛ طµط­ ط§ط¶ط؛ط· ط­ظپط¸', 'if the amount is correct, tap Save'],
  ['ط¯ظ‡ ط¥ظٹطµط§ظ„ ط£ظ†ظ‡ظٹ ظ…ط§ظƒظٹظ†ط©طں', 'Which machine is this receipt for?'],
  ['ط§ظ„ظ…ظ„ظپ ط§ظ„ط£طµظ„ظٹ ظ…ط´ ظ…ظˆط¬ظˆط¯', 'Original file not found'],
  ['ظ…ط§ظƒظٹظ†ط§طھ ظ…ط§ ظˆطµظ„ط´ ظ…ظ„ظپظ‡ط§', 'Machines with no file received'],
  ['طھط¹ط¯ظٹظ„ ط¨ظٹط§ظ†ط§طھ ط§ظ„ط¥ظٹطµط§ظ„', 'Edit receipt data'],
  ['ط±ط§ط¬ط¹ظ‡ ظ…ظ† ط²ط± ط§ظ„ظ…ط±ط§ط¬ط¹ط©', 'review it using the review button'],
  ['ط§ظ„ظ…طµط¯ط±: ط±ط¬ظˆط¹ ظ„ظˆط§طھط³ط§ط¨', 'Source: back to WhatsApp'],
  ['ط¨ط¯ط، ط§ظ„ظ…ط¹ط§ظ„ط¬ط© ط§ظ„ط´ط§ظ…ظ„ط©', 'Run full processing'],
  ['ط¨ظٹظ†ظ…ط§ ط£ط؛ظ„ط¨ ط§ظ„ط¥ظٹطµط§ظ„ط§طھ', 'while most receipts are'],
  ['ط§ظپطھط­ ط§ظ„ظپظˆظ„ط¯ط± ط¨ظˆط§ط³ط·ط©', 'Open folder with'],
  ['ط§ظ„ظ…ط¨ظ„ط؛ ظٹط­طھط§ط¬ ظ…ط±ط§ط¬ط¹ط©', 'Amount needs review'],
  ['ط¥ظٹطµط§ظ„ ظ…ظƒط±ط± ظ„ظ„ظ…ط§ظƒظٹظ†ط©', 'Duplicate receipt for machine'],
  ['ظ…ظ‚ط¯ط±طھط´ ط£ط¯ظˆظ‘ط± ط§ظ„طµظˆط±ط©', 'Could not rotate the image'],
  ['ط§طھط£ظƒط¯ ظ…ظ†ظ‡ ظˆط§ط¶ط؛ط· ط­ظپط¸', 'verify it and tap Save'],
  ['ظ„ظ„ظ…ط±ط§ط¬ط¹ط© (ظ…ط´ ظ…ط¹ط¯ظˆط¯)', 'For review (not counted)'],
  ['ط±ط§ط¬ط¹ظ‡ط§ ظ„ظˆ ط§ظ„ط¹ط¯ط¯ ظ…ظ‡ظ…', 'check them if the count matters'],
  ['ط§ظ„ط¹ط¯: ط¯ظ‚ظٹظ‚ â†گ ط§ظ„ط³ط±ظٹط¹', 'Counting: precise, then fast'],
  ['ظ…طھطµظپط­ ظ…ظ„ظپط§طھ ط§ظ„ظ†ط¸ط§ظ…', 'System file browser'],
  ['ط­ظپط¸ ظˆطھط­ط¯ظٹط« ط§ظ„ظ…ظ„ظپط§طھ', 'Save and update files'],
  ['ط¨ظ‚ط±ط£ طµظپط­ط§طھ ظ…ظ„ظپ PDF', 'Reading pages of PDF file'],
  ['ط±ط§ط¬ط¹ ط§ظ„ظ…ط¨ظ„ط؛ ظٹط¯ظˆظٹط§ظ‹', 'check the amount manually'],
  ['ط§ظ„ط¥ط®ط±ط§ط¬: ط§ظ„ط§ظپطھط±ط§ط¶ظٹ', 'Output: default'],
  ['ط§ط®طھظٹط§ط± ظ‡ط°ط§ ط§ظ„ظپظˆظ„ط¯ط±', 'Choose this folder'],
  ['ظ…ط³طھط¨ط¹ط¯ (ط¹ظ„ط§ظ…ط© ط؛ظ„ط·)', 'Excluded (marked wrong)'],
  ['ط¹ط¯ط¯ ط¥ظٹطµط§ظ„ط§طھ ط§ظ„ط¯ظپط¹:', 'Payment receipts count:'],
  ['ط¥ظ„ط؛ط§ط، ط¹ظ„ط§ظ…ط© طµط­/ط؛ظ„ط·', 'Clear correct/wrong mark'],
  ['ظ…ط´ ظ…ظˆط¬ظˆط¯ ظپظٹ ظپظˆظ„ط¯ط±', 'not found in folder'],
  ['ط§ظ„ظˆط­ط¯ط© ط؛ظٹط± ظ…ط¹ط±ظˆظپط©', 'Unknown unit'],
  ['ظپطھط­ ظپظˆظ„ط¯ط± ط§ظ„ط¥ط®ط±ط§ط¬', 'Open output folder'],
  ['ط§ظ„طھط§ط±ظٹط® طµط­ - ط³ظٹط¨ظ‡', 'The date is correct - keep it'],
  ['ظ…ط´ طھط¨ط¹ ظ†ظپط³ ط§ظ„ظٹظˆظ…:', 'Not from the same day:'],
  ['ط¨ط³ ط§ظ„طھط³ظ…ظٹط© ظ…ط§ظƒظٹظ†ط©', 'but the naming says machine'],
  ['ظ„ظˆط¬ ط¥ظٹطµط§ظ„ط§طھ ط§ظ„ط¯ظپط¹', 'Payment receipts log'],
  ['ط§ظ„ط£ظٹط§ظ… ط§ظ„ظ„ظٹ ظپط§طھطھ', 'previous days'],
  ['طھظ… طھظˆظ„ظٹط¯ ط§ظ„ط¥ظƒط³ظٹظ„', 'Excel generated'],
  ['ظ…ظ„ط§ط­ط¸ط§طھ ظ„ظ„ظ…ط±ط§ط¬ط¹ط©', 'Notes for review'],
  ['ط§ظ„ظ…ط¨ظ„ط؛ ط؛ظٹط± ظ…ظ‚ط±ظˆط،', 'Amount unreadable'],
  ['ظ…ط´ ط¨ط§طھط´ - طھط¬ط§ظ‡ظ„ظ‡', 'Not a batch - ignore it'],
  ['ط§ظ„ظ‚ط±ط§ط،ط© ط§ظ„طھط§ظ†ظٹط©:', 'Second reading:'],
  ['ط£ط±ظ‚ط§ظ… ظ…ط´ ظ„ط§ظ‚ظٹظ‡ط§:', 'Numbers not found:'],
  ['ظ…ظ„ظپ ط¹ظ„ظ‰ ط¥ظ†ظ‡ظ… طµط­طں', 'file(s) as correct?'],
  ['ظ…ظ„ظپ ظˆظ†طھظٹط¬ط© ط§ظ„ظپط­طµ', 'file(s) and the scan result'],
  ['ط§طھط£ظƒط¯ ظ…ظ† ط§ظ„ظ…ط¨ظ„ط؛', 'check the amount'],
  ['ظ…ظپظٹط´ ط­ط§ط¬ط© ظ„ظ„ظ…ط³ط­', 'Nothing to delete'],
  ['ط§ظ„ظ…ط¨ظ„ط؛ ط§ظ„ط¥ط¬ظ…ط§ظ„ظٹ', 'Total amount'],
  ['ط§ظ„ط¥ط¬ظ…ط§ظ„ظٹ ط§ظ„ط¹ط§ظ…:', 'Grand total:'],
  ['ط§ظ„ظ…طµط¯ط±: ظ…ظ„ظپ PDF', 'Source: PDF file'],
  ['ظˆط§طھط³ط§ط¨ (طھظ„ظ‚ط§ط¦ظٹ)', 'WhatsApp (auto)'],
  ['ظˆط­ط¯ط© ط؛ظٹط± ظ…ط¹ط±ظˆظپط©', 'Unknown unit'],
  ['ط¯ظ‡ ط¨ط§طھط´ - ط§ط¯ط®ظ„ظ‡', 'This is a batch - include it'],
  ['ط§طھط¹ط±ظپطھ ظƒظ…ط§ظ† ظپظٹ:', 'also identified in:'],
  ['ط§ظ„ط¹ط¯ ظ„ظ‚ظ‰ ظ…ط§ظƒظٹظ†ط©', 'Counting found machine'],
  ['ط§ظ„ظ…ظ„ظپ ظ…ط´ ظ…ظˆط¬ظˆط¯', 'File not found'],
  ['ظ…ظ„ظپ ظ…ظ† ط§ظ„ط¥ط®ط±ط§ط¬', 'file(s) from the output'],
  ['طھظ‚ط±ظٹط± ط§ظ„ظ…ط±ط§ط¬ط¹ط©', 'Review report'],
  ['ظپط´ظ„ ظ…ظ„ط، ط§ظ„ظ‚ط§ظ„ط¨', 'Failed to fill the template'],
  ['ظ…ط§ظƒظٹظ†ط§طھ ظ†ط§ظ‚طµط©:', 'Missing machines:'],
  ['ظ…ط´ظƒظˆظƒ ط¥ظ†ظ‡ ظ…ظƒط±ط±', 'possible duplicate'],
  ['ط±ط¬ظ‘ط¹ظ‡ (ط§ط­طھط³ط¨ظ‡)', 'Restore (count it)'],
  ['ط§ظƒطھط¨ ظ…ط¨ظ„ط؛ طµط­ظٹط­', 'Enter a valid amount'],
  ['ظ…ظƒط±ط± - ط§ط³طھط¨ط¹ط¯ظ‡', 'Duplicate - exclude it'],
  ['ط§ط®طھط§ط± ط§ظ„ظ…ط§ظƒظٹظ†ط©', 'Select machine'],
  ['ط§طھط³ظ…ظ‘ظ‰طŒ ط¥ط¬ظ…ط§ظ„ظٹ', 'named, total'],
  ['طھط­ط¯ظٹط¯ ط§ظ„ظ…ط§ظƒظٹظ†ط©', 'Set machine'],
  ['ط¥ظٹطµط§ظ„ط§طھ ط§ظ„ط¯ظپط¹', 'Payment receipts'],
  ['ط¥ط¬ظ…ط§ظ„ظٹ ظ…ط§ظƒظٹظ†ط©', 'Total of machine'],
  ['ط¨ط¹ظٹط¯ ط¹ظ† ظ…طھظˆط³ط·', 'is far from the average of'],
  ['ظ…ط³طھط¨ط¹ط¯ (ظ…ظƒط±ط±)', 'Excluded (duplicate)'],
  ['ط±ظ‚ظ… ط§ظ„ظ…ط§ظƒظٹظ†ط©:', 'Machine number:'],
  ['ظ…ط§ظƒظٹظ†ط§طھ ظ†ط§ظ‚طµط©', 'Missing machines'],
  ['طھط§ط±ظٹط® ط§ظ„ط¥ظٹطµط§ظ„', 'Receipt date'],
  ['ط§طھط³ظ…ظ‘ظ‰ ظ…ط§ظƒظٹظ†ط©', 'Named machine'],
  ['ظ…ظ„ظپ ظ…ط³ظ…ظ‘ظ‰ ظ…ظ†:', 'named file(s) from:'],
  ['ط§طھط³ظ…ظ‘ظ‰ ظˆط§طھظ†ط³ط®', 'named and copied'],
  ['طھط¹ط±ظٹظپ ط§ظ„ظ…ظ„ظپط§طھ', 'Identifying files'],
  ['ط¹ط¯ظ‘ ط§ظ„ط¥ظٹطµط§ظ„ط§طھ', 'Counting receipts'],
  ['ظ‚ط±ط§ط،ط© طھط§ظ†ظٹط©:', 'Second reading:'],
  ['ظپط´ظ„طھ ط§ظ„ظ‚ط±ط§ط،ط©', 'read failed'],
  ['ظ†ظپط³ ط§ظ„ظ…ط§ظƒظٹظ†ط©', 'Same machine'],
  ['ظ…ط³طھظ†ط¯ ظ…طھط¬ط§ظ‡ظ„', 'Ignored document'],
  ['ط§ط®طھظٹط§ط± ظپظˆظ„ط¯ط±', 'Choose folder'],
  ['ظپط´ظ„ ط§ظ„طھط³ظ…ظٹط©:', 'Naming failed:'],
  ['طھظ…طھ ط§ظ„طھط³ظ…ظٹط©:', 'Naming done:'],
  ['ظپط´ظ„ ط§ظ„طھط¹ط±ظٹظپ:', 'Identification failed:'],
  ['ظ†ط§ظ‚طµ ط§ظ„ظ‚ط±ط§ط،ط©', 'partial read'],
  ['ظ„ظˆط¬ ط§ظ„طھظ‚ظپظٹظ„', 'Closing log'],
  ['طھط§ط±ظٹط® ظ…ط®طھظ„ظپ', 'Different date'],
  ['ط§ظ†طھظ‡ظ‰ ط§ظ„ظپط­طµ', 'Scan finished'],
  ['ظˆظ†ظپط³ ط§ظ„ظ…ط¨ظ„ط؛', 'and the same amount'],
  ['ط±ط§ط¬ط¹ ط§ظ„طµظˆط±ط©', 'check the image'],
  ['ط®ظ„ظٹظ‡ ط¨طھط§ط±ظٹط®', 'Set it to date'],
  ['ط­ظپط¸ ظˆط§ظ„طھط§ظ„ظٹ', 'Save & next'],
  ['ط§طھظˆظ‚ظپ. ظپط§ط¶ظ„', 'Stopped. Remaining'],
  ['ط§ظƒطھظ…ظ„ ط§ظ„ط¹ط¯:', 'Counting complete:'],
  ['ط¥ط±ط³ط§ظ„ ط§ظ„ظ…ظ„ظپ', 'Send file'],
  ['طھط؛ظٹظٹط± ط§ظ„ظ„ط؛ط©', 'Change language'],
  ['ط§ظ„ط¥ط³ظƒظ†ط¯ط±ظٹط©', 'Alexandria'],
  ['ظ…ظ„ط، ط§ظ„ط´ط§ط´ط©', 'Full screen'],
  ['ط¨ط³ ط§ظ„طھط³ظ„ط³ظ„', 'but the sequence'],
  ['ط¢ط®ط± ظ†طھظٹط¬ط©:', 'Last result:'],
  ['ظپط­طµ ظˆطھط³ظ…ظٹط©', 'Scan & name'],
  ['ط¥ط±ط³ط§ظ„ ط§ظ„ظƒظ„', 'Send all'],
  ['ط§طھط±ط§ط¬ط¹: طµط­', 'Reviewed: correct'],
  ['ط¥ط¹ط§ط¯ط© ط§ظ„ط¹ط¯', 'Recount'],
  ['ظ…ط®طھظ„ظپط© ط¹ظ†', 'differs from'],
  ['ط§طھط£ظƒط¯ ظ…ظ†ظ‡', 'verify it'],
  ['ط؛ظٹط± ظ…ظ‚ط±ظˆط،', 'unreadable'],
  ['ط§ظ„ط¨ط§ظ‚ظٹ طµط­', 'The rest are correct'],
  ['ظ…ط³ط­ ط§ظ„ظƒظ„', 'Delete all'],
  ['ط؛ظٹط± ظ…ط­ط¯ط¯', 'Unspecified'],
  ['ظ„ظ„ظ…ط±ط§ط¬ط¹ط©', 'for review'],
  ['ط§ظ„ظ…ط§ظƒظٹظ†ط©', 'Machine'],
  ['ط§ظ„ظ…ظ‚ط±ظˆط،:', 'Read:'],
  ['ط§ط³طھط®ط¯ظ…ظ‡ط§', 'Use it'],
  ['ظپط´ظ„ ط§ظ„ط¹ط¯', 'Counting failed'],
  ['ط¹ط¯ظ‘ ط³ط±ظٹط¹', 'Quick count'],
  ['ط¨ط¯ط، ط§ظ„ط¹ط¯', 'Start counting'],
  ['ط§ظ„طھظ‚ظپظٹظ„', 'Closing'],
  ['ظ…ط§ظƒظٹظ†ط§طھ', 'Machines'],
  ['ظˆط§ظ„ط£ط؛ظ„ط¨', 'while most are'],
  ['ط§ظ„ظ…ط¨ظ„ط؛:', 'Amount:'],
  ['ط§ظ„ط¥طµط¯ط§ط±', 'Version'],
  ['ط§ظپطھط±ط§ط¶ظٹ', 'default'],
  ['ط§ظ„طھط³ظ„ط³ظ„', 'Sequence'],
  ['ط§ظ„ظپطھط±ط©', 'Period'],
  ['ظ‡طھطھظ…ط³ط­', 'Will delete'],
  ['طھظ… ظ…ط³ط­', 'Deleted'],
  ['ظ…ط§ظƒظٹظ†ط©', 'Machine'],
  ['ط§ظƒطھظ…ظ„!', 'Done!'],
  ['ط¨ط§طھط´ط§طھ', 'batches'],
  ['طھط§ط±ظٹط®ظ‡', 'Date'],
  ['ظ…ط³طھط¨ط¹ط¯', 'Excluded'],
  ['طھظ… ظپط­طµ', 'Scanned'],
  ['ط¨طھط§ط±ظٹط®', 'dated'],
  ['ظ‡طھط¹ظ„ظ‘ظ…', 'Mark'],
  ['ظ‡ظٹطھظ…ط³ط­', 'Will delete'],
  ['ط§ظ„ط£طµظ„:', 'Original:'],
  ['ط§طھط³ظ…ظ‘ظ‰', 'named'],
  ['ظ…ط±ط§ط¬ط¹ط©', 'Review'],
  ['ظ…ظ„ظپط§طھ', 'files'],
  ['ط¥ظٹطµط§ظ„', 'receipt(s)'],
  ['طµظپط­ط§طھ', 'pages'],
  ['ط§ظ„ط¹ط¯ط¯', 'Count'],
  ['ظ…ظ‚ط±ظˆط،', 'read'],
  ['ط£ظٹط§ظ…', 'days'],
  ['ظ‚ط§ظ„ط¨', 'Template'],
  ['ط¨ط§طھط´', 'batch'],
  ['ظ†ط§ظ‚طµ', 'missing'],
  ['طµظپط­ط©', 'page(s)'],
  ['طµظˆط±ط©', 'image(s)'],
  ['ظپط§ط¶ظ„', 'remaining'],
  ['ط±ط§ط¬ط¹', 'review'],
  ['ط§ظ„ظƒظ„', 'All'],
  ['ظƒط§ظ…ظ„', 'full'],
  ['ظ…ظ„ظپ', 'file(s)'],
  ['ط¬.ظ…', 'EGP'],
  ['ظپط­طµ', 'Scan'],
  ['ظپظٹظ‡', 'There are'],
  ['ظ„ط³ظ‡', 'yet'],
  ['ط£ظˆ', 'or'],
  ['ظˆ', 'and']
];

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ ط£ظ„ظˆط§ظ† طھطھط£ظ‚ظ„ظ… ظ…ط¹ ط§ظ„ط¯ط§ط±ظƒ ظ…ظˆط¯ â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
bool _isDark(BuildContext c) => Theme.of(c).brightness == Brightness.dark;

Color _tintBg(BuildContext c, MaterialColor m, {bool soft = false, bool strong = false}) {
  if (_isDark(c)) return m.shade900.withAlpha(strong ? 210 : 140);
  return strong ? m.shade300 : (soft ? m.shade50 : m.shade100);
}

Color _tintBorder(BuildContext c, MaterialColor m) => _isDark(c) ? m.shade300 : m.shade800;
Color _tintFg(BuildContext c, MaterialColor m) => _isDark(c) ? m.shade200 : m.shade900;

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ ط§ظ„ظ‡ظٹط¯ط±: ط§ط³ظ… ط§ظ„طھط·ط¨ظٹظ‚ + ط²ط±ط§ظٹط± ط§ظ„ظ„ط؛ط© ظˆط§ظ„ط¯ط§ط±ظƒ â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

// ط¹ظ„ظ‰ ط§ظ„ط´ط§ط´ط§طھ ط§ظ„ط¹ط±ظٹط¶ط©: ط£ظٹظ‚ظˆظ†ط© ظ…ظ†ظپطµظ„ط©. ط¹ظ„ظ‰ ط§ظ„طµط؛ظٹط±ط©: ط¨طھطھظ†ظ‚ظ„ ظ„ظ‚ط§ظٹظ…ط© "ط§ظ„ظ…ط²ظٹط¯" (ط¨ظ†ط¯ "ط§ظ„ظ…ط¸ظ‡ط± ظˆط§ظ„ظ„ط؛ط©").
class HeaderButtons extends StatelessWidget {
  const HeaderButtons({super.key});

  @override
  Widget build(BuildContext context) {
    if (MediaQuery.sizeOf(context).width < 400) return const SizedBox.shrink();
    final ar = AppSettings.of(context).isAr;
    return IconButton(
      tooltip: ar ? 'ط§ظ„ظ…ط¸ظ‡ط± ظˆط§ظ„ظ„ط؛ط©' : 'Theme & language',
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
            label(ar ? 'ط§ظ„ظ„ط؛ط©' : 'Language'),
            SegmentedButton<String>(
              showSelectedIcon: false,
              segments: [
                ButtonSegment<String>(value: 'sys', label: w.Text(ar ? 'طھظ„ظ‚ط§ط¦ظٹ' : 'Auto')),
                const ButtonSegment<String>(value: 'ar', label: w.Text('ط§ظ„ط¹ط±ط¨ظٹط©')),
                const ButtonSegment<String>(value: 'en', label: w.Text('English')),
              ],
              selected: {s.langCode ?? 'sys'},
              onSelectionChanged: (v) => s.setLang(v.first == 'sys' ? null : v.first),
            ),
            label(ar ? 'ط§ظ„ظ…ط¸ظ‡ط±' : 'Theme'),
            SegmentedButton<ThemeMode>(
              showSelectedIcon: false,
              segments: [
                ButtonSegment<ThemeMode>(
                  value: ThemeMode.system,
                  icon: const Icon(Icons.brightness_auto, size: 18),
                  label: w.Text(ar ? 'طھظ„ظ‚ط§ط¦ظٹ' : 'Auto'),
                ),
                ButtonSegment<ThemeMode>(
                  value: ThemeMode.light,
                  icon: const Icon(Icons.light_mode, size: 18),
                  label: w.Text(ar ? 'ظپط§طھط­' : 'Light'),
                ),
                ButtonSegment<ThemeMode>(
                  value: ThemeMode.dark,
                  icon: const Icon(Icons.dark_mode, size: 18),
                  label: w.Text(ar ? 'ط¯ط§ظƒظ†' : 'Dark'),
                ),
              ],
              selected: {s.themeMode},
              onSelectionChanged: (v) => s.setMode(v.first),
            ),
            label(ar ? 'ظ„ظˆظ† ط§ظ„طھط·ط¨ظٹظ‚' : 'Accent color'),
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

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ ط¬ط¯ط§ظˆظ„ ط§ظ„ظ…ط§ظƒظٹظ†ط§طھ ظˆط§ظ„ظˆط­ط¯ط§طھ â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
class TerminalInfo {
  final int orderNum;
  final int excelRow;
  final String unitName;
  const TerminalInfo(this.orderNum, this.excelRow, this.unitName);
}

// ط§ظ„ظ†ط³ط®ط© ط§ظ„ط¹ط§ظ…ط©: ظ…ظپظٹط´ ظ…ط§ظƒظٹظ†ط§طھ ظˆظ„ط§ ظˆط­ط¯ط§طھ ظ…ظƒطھظˆط¨ط© ظپظٹ ط§ظ„ظƒظˆط¯. ظƒظ„ظ‡ ط¨ظٹطھظ‚ط±ط£ ظ…ظ† ظ…ظ„ظپ A2R_Settings.xlsx
final Map<String, TerminalInfo> terminals = {};
final Map<int, String> numToTid = {};
final Map<String, String> tidFix = {};
final Map<String, String> tidAlias = {};
final Map<String, int> merchantCodeMap = {};
final List<MapEntry<List<String>, int>> nameHints = [];
final Map<int, String> rowNames = {};

// ط¨ظٹط¨ظ†ظٹ ط§ظ„ظپظ‡ط§ط±ط³ (ط±ظ‚ظ… ط§ظ„ظ…ط§ظƒظٹظ†ط© -> TID طŒ ط±ظ‚ظ… ط§ظ„طµظپ -> ط§ط³ظ… ط§ظ„ظˆط­ط¯ط©) ظ…ظ† ط§ظ„ظ…ط§ظƒظٹظ†ط§طھ ط§ظ„ظ…ط­ظ…ظ‘ظ„ط©
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
  return 'ظ…ط§ظƒظٹظ†ط© $i${t == null || t.unitName.isEmpty ? '' : ' - ${t.unitName}'}';
}

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ ط®ظˆط§ط±ط²ظ…ظٹط§طھ ط§ظ„ط§ط³طھط®ط±ط§ط¬ (ظ…ظ†ظ‚ظˆظ„ط© ظ…ظ† ظ†ط³ط®ط© ط§ظ„ط¨ط§ظٹط«ظˆظ†) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

// ط³ط·ط± ظ†طµ ظ…ط¹ ظ…ظƒط§ظ†ظ‡ ظپظٹ ط§ظ„طµظˆط±ط©
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

// ظٹط±طھط¨ ط§ظ„ظ†طµ طµظپظˆظپ (ط§ظ„ط§ط³ظ… ظٹظ…ظٹظ† ظˆط§ظ„ط±ظ‚ظ… ط¬ظ†ط¨ظ‡ ظپظٹ ظ†ظپط³ ط§ظ„ط³ط·ط±) ط²ظٹ Tesseract ط¨ط§ظ„ط¸ط¨ط·
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

// طµظˆط±ط© ظپظٹظ‡ط§ ط£ظƒطھط± ظ…ظ† ط¥ظٹطµط§ظ„ ط¬ظ†ط¨ ط¨ط¹ط¶ -> ظ†ظ‚ط³ظ…ظ‡ط§ ظ„ط£ط¹ظ…ط¯ط©
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

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ ط§ظ„ظ…ط¨ط§ظ„ط؛ â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

class Decision {
  final double val;
  final bool ok;
  Decision(this.val, this.ok);
}

Decision decide(List<Parsed> results) {
  // 1) طھط·ط§ط¨ظ‚ طھط§ظ… ط¨ظٹظ† ظ…ط¬ظ…ظˆط¹ ط§ظ„ظƒط±ظˆطھ ظˆ GRAND TOTALS
  for (final r in results) {
    final b = r.bsum;
    if (b != null && b > 0 && r.grand.any((g) => (g - b).abs() < 0.005)) {
      return Decision(b, true);
    }
  }
  // 2) ظپط±ظ‚ ط±ظ‚ظ… ظˆط§ط­ط¯
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
  // 3) ط¥ط¬ظ…ط§ظ„ظٹ ظ…ظƒط±ط±
  for (final r in results) {
    final c = mostCommon(r.grand);
    if (c != null && c.value >= 2 && r.bsum == null) return Decision(c.key, false);
  }
  // 4) طھطµظˆظٹطھ
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

// ط¥ظٹطµط§ظ„ ط¯ظپط¹ (RRN / ط±ظ‚ظ… ظƒط§ط±طھ ظ…ط®ظپظٹ / Approved ...) ظˆظ„ط§ ط¨ط§طھط´ ظƒظ„ظˆط² / ظ…ظ„ط®طµ / ط±ظٹ ط¨ط±ظ†طھطں
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

// ظ„ظˆ GRAND TOTAL = ظ…ط¬ظ…ظˆط¹ ط¥ط¬ظ…ط§ظ„ظٹط§طھ ط§ظ„ظƒط±ظˆطھ (ط­طھظ‰ ظ„ظˆ ط£ط³ظ…ط§ط، ط§ظ„ظƒط±ظˆطھ ط§طھظ‚ط±طھ ط؛ظ„ط·): ظ†ط±ط¬ظ‘ط¹ظ‡ ظ…ط¤ظƒط¯.
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

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ ط§ظ„ظ…ط§ظƒظٹظ†ط© / ط§ظ„ظˆط­ط¯ط© / ط§ظ„طھط§ط±ظٹط® â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
String? _tidLabeledRaw(String t) {
  final m = RegExp(r"TERMINAL\s*(?:ID)?\D{0,15}?(\d[\d\s/]{6,9}\d)",
          caseSensitive: false)
      .firstMatch(t);
  if (m == null) return null;
  final raw = m.group(1)!.replaceAll('/', '7').replaceAll(RegExp(r'\D'), '');
  if (raw.length < 8) return null;
  final d = raw.substring(raw.length - 8);
  return tidFix[d] ?? d;
}

String? _tidLabeled(String t) {
  final d = _tidLabeledRaw(t);
  return d != null && terminals.containsKey(d) ? d : null;
}

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ طھط¹ظ„ظ‘ظ… ط§ظ„ظ…ط§ظƒظٹظ†ط§طھ ظˆط§ظ„ظˆط­ط¯ط§طھ ظ…ظ† ط§ظ„ط¥ظٹطµط§ظ„ط§طھ â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

// ط§ط³ظ… ط§ظ„ظˆط­ط¯ط© ط§ظ„ظ…ظƒطھظˆط¨ ط¹ظ„ظ‰ ط§ظ„ط¥ظٹطµط§ظ„: ط¨ط¹ط¯ ظƒظ„ظ…ط© MERCHANT NAME ظ„ظˆ ظ…ظˆط¬ظˆط¯ط©طŒ ظˆط¥ظ„ط§ ط£ظˆظ„ ط³ط·ط± ط§ط³ظ… ظپظٹ ط£ط¹ظ„ظ‰ ط§ظ„ط¥ظٹطµط§ظ„
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
  final skip = RegExp(
      r'NATIONAL|BANK|\bNBE\b|EGYPT|RECEIPT|CUSTOMER|COPY|MERCHANT|TERMINAL|BATCH|DATE|TIME|SALE|TOTAL|AMOUNT|VISA|MASTER|MEEZ|THANK|SUCCESS|CLOSED|CARD|APPROVAL|\bRRN\b|TRACE|VERSION|GRAND|REPORT|SETTLEMENT',
      caseSensitive: false);
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

// ظ…ط§ظƒظٹظ†ط© ط¬ط¯ظٹط¯ط© ظ…ظƒطھظˆط¨ط© ط¹ظ„ظ‰ ط§ظ„ط¥ظٹطµط§ظ„ (Terminal ID) ظˆظ…ط´ ظپظٹ ط§ظ„ظ‚ط§ظٹظ…ط©: ط¨ظ†ط³ط¬ظ‘ظ„ظ‡ط§ ظ‡ظٹ ظˆظˆط­ط¯طھظ‡ط§ ظˆط¨ظ†ط­ظپط¸ظ‡ط§ ظپظٹ ظ…ظ„ظپ ط§ظ„ط¥ط¹ط¯ط§ط¯ط§طھ
String? discoverTerminal(List<String> texts) {
  final counts = <String, int>{};
  for (final t in texts) {
    final d = _tidLabeledRaw(t);
    if (d != null) counts[d] = (counts[d] ?? 0) + 1;
  }
  if (counts.isEmpty) return null;
  final tid = counts.entries.reduce((a, b) => a.value >= b.value ? a : b).key;
  if (terminals.containsKey(tid)) return tid;
  final name = extractUnitName(texts) ?? 'Unit $tid';
  final row = _rowForUnitName(name);
  var order = 0;
  for (final t in terminals.values) {
    if (t.orderNum > order) order = t.orderNum;
  }
  terminals[tid] = TerminalInfo(order + 1, row, rowNames[row] ?? name);
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

// ط¹ط¯ط¯ ظ…ط¹ط§ظ…ظ„ط§طھ ط§ظ„ط¨ط§طھط´ ظ…ظ† ط¥ظٹطµط§ظ„ ط§ظ„ط¥ظ‚ظپط§ظ„: "GRAND TOTALS / Sale (010)" (ظˆظ„ظˆ ط§طھظ‚ط±ط§ ظ…ط¬ظ…ظˆط¹ ط§ظ„ظƒط±ظˆطھ ط¨ظٹطھط·ط§ط¨ظ‚ ظ…ط¹ط§ظ‡)
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
  if (g != null && s != null && g != s) return null; // ط§ظ„ظ‚ط±ط§ط،طھظٹظ† ظ…طھط¹ط§ط±ط¶طھظٹظ†: ط§ظ„ط£ظپط¶ظ„ طھظƒطھط¨ظ‡ ط¨ط¥ظٹط¯ظƒ
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
  int? txn; // ط¹ط¯ط¯ ظ…ط¹ط§ظ…ظ„ط§طھ ط§ظ„ط¨ط§طھط´ (ظ…ظ† ط¥ظٹطµط§ظ„ ط§ظ„ط¥ظ‚ظپط§ظ„)
  Analysis(this.tid, this.how, this.row, this.unitName, this.val, this.ok,
      this.batches, this.date);
}

// طھط­ظ„ظٹظ„ طµظˆط±ط© ظƒط§ظ…ظ„ط© (ظ…ظ…ظƒظ† ظپظٹظ‡ط§ ط£ظƒطھط± ظ…ظ† ط¥ظٹطµط§ظ„)
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
    if (b != null && seen.contains(key)) continue; // ظ†ظپط³ ط§ظ„ط¨ط§طھط´ ط§طھطµظˆط± ظ…ط±طھظٹظ†
    seen.add(key);
    batches.add(r.val);
    txns.add(batchTxnOf(r.texts));
    allOk = allOk && r.ok;
  }
  var total = round2(batches.fold(0.0, (a, b) => a + b));
  var ok = allOk && batches.isNotEmpty;

  // ط´ط¨ظƒط© ط£ظ…ط§ظ†: ظ„ظˆ ط§ظ„طھظ‚ط³ظٹظ… ظ…ط´ ظ…ط¤ظƒط¯ ط¬ط±ظ‘ط¨ ط§ظ„طµظˆط±ط© ظƒظ„ظ‡ط§
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
  result.txn = (txns.isNotEmpty && txns.every((e) => e != null)) ? txns.fold<int>(0, (a, e) => a + e!) : null;
  return result;
}

double round2(double val) => double.parse(val.toStringAsFixed(2));

// ط¯ظ…ط¬ ط§ظ„ظ‚ط±ط§ط،ط© ط§ظ„طھط§ظ†ظٹط© (ط§ظ„طµظˆط±ط© ط§ظ„ظ…ط­ط³ظ‘ظ†ط©) ظپظٹ ط§ظ„ظ‚ط±ط§ط،ط© ط§ظ„ط£ظˆظ„ظ‰.
// ط¨طھط±ط¬ط¹ ظ‚ظٹظ…ط© ط§ظ„ظ…ط¨ظ„ط؛ ط§ظ„طھط§ظ†ظٹط© ظ„ظˆ ظ…طھط¹ط§ط±ط¶ط© ظ…ط¹ ط§ظ„ط£ظˆظ„ظ‰ (ط¹ظ„ط´ط§ظ† طھطھط±ط§ط¬ط¹)طŒ ظˆط¥ظ„ط§ null.
double? mergeSecondRead(Analysis a, Analysis b) {
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
  if ((a.val - b.val).abs() < 0.005) return null; // ط§ظ„ظ‚ط±ط§ط،طھظٹظ† ظ…طھط·ط§ط¨ظ‚طھظٹظ†
  if (!a.ok && b.ok) {
    // ط§ظ„ط£ظˆظ„ظ‰ ظ…ط´ ظ…ط¤ظƒط¯ط© ظˆط§ظ„طھط§ظ†ظٹط© ظ…ط¤ظƒط¯ط© ط¨ظ…ط¬ظ…ظˆط¹ ط§ظ„ظƒط±ظˆطھ -> ظ†ط¹طھظ…ط¯ ط§ظ„طھط§ظ†ظٹط©
    a.val = b.val;
    a.ok = true;
    a.batches = [...b.batches];
    a.txn = b.txn ?? a.txn;
    return null;
  }
  if (a.ok && !b.ok) return null; // ط§ظ„ط£ظˆظ„ظ‰ ظ…ط¤ظƒط¯ط©طŒ ظ†طھط¬ط§ظ‡ظ„ ط§ظ„طھط§ظ†ظٹط©
  a.ok = false; // ظ‚ط±ط§ط،طھظٹظ† ظ…ط®طھظ„ظپطھظٹظ†: ظ„ط§ط²ظ… ظ…ط±ط§ط¬ط¹ط©
  return b.val;
}

// ط¨ظٹط­ظˆظ‘ظ„ طµظپط­ط§طھ ظ…ظ„ظپ PDF ظ„طµظˆط± PNG (ط¨ظٹط³طھط®ط¯ظ… ظپظٹ ط­ط§ظ„ط© ظ…ظپظٹط´ طµظˆط±)
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

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ ظ…ط³ط§ط±ط§طھ ط§ظ„ظ…ط¬ظ„ط¯ط§طھ â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const String root = '/storage/emulated/0';
final String baseDir = '$root/A2R-POS-Public';
final String rawImagesDir = '$baseDir/raw_images';
String outputDir = '$baseDir'; // ظ…ظ…ظƒظ† ط§ظ„ظ…ط³طھط®ط¯ظ… ظٹط؛ظٹظ‘ط±ظ‡ ظ…ظ† ط§ظ„ط´ط§ط´ط©
final String defaultOutputDir = '$baseDir';
final String tmpDir = '$baseDir/.tmp';
final String cfgPath = '$baseDir/.app_cfg.json';
final String settingsPath = '$baseDir/A2R_Settings.xlsx';

String baseName(String path) => path.split('/').last;

List<Directory> _subDirs(String path) {
  try {
    final dir = Directory(path);
    if (!dir.existsSync()) return [];
    return dir.listSync().whereType<Directory>().toList();
  } catch (_) {
    return [];
  }
}

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
    for (final pkg in _subDirs(mediaDir.path)) {
      if (baseName(pkg.path).toLowerCase().contains('whatsapp')) {
        checkAndAdd('${pkg.path}/Media/WhatsApp Images');
        checkAndAdd('${pkg.path}/WhatsApp/Media/WhatsApp Images');
        for (final acc in _subDirs('${pkg.path}/WhatsApp/accounts')) {
          checkAndAdd('${acc.path}/Media/WhatsApp Images');
        }
        for (final acc in _subDirs('${pkg.path}/accounts')) {
          checkAndAdd('${acc.path}/Media/WhatsApp Images');
        }
      }
    }
  }

  for (final d in _subDirs(root)) {
    if (baseName(d.path).toLowerCase().startsWith('whatsapp')) {
      checkAndAdd('${d.path}/Media/WhatsApp Images');
    }
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

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ ظ…ط¹ط§ظ„ط¬ط© ط§ظ„طµظˆط± (طھط¯ظˆظٹط± / طھط­ط³ظٹظ†) - ط¨طھط´طھط؛ظ„ ظپظٹ Isolate â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// ط¨طھطµظ„ظ‘ط­ ط§طھط¬ط§ظ‡ ط§ظ„ظ€ EXIF ظˆط¨طھط¯ظˆظ‘ط± ط§ظ„طµظˆط±ط© ظ„ظˆ deg != 0. ط¨طھط±ط¬ط¹ true ظ„ظˆ ظƒطھط¨طھ ظ…ظ„ظپ ط¬ط¯ظٹط¯.
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

// ظ†ط³ط®ط© ظ…ط­ط³ظ‘ظ†ط© ظ„ظ„ظ‚ط±ط§ط،ط© ط§ظ„طھط§ظ†ظٹط©: ط±ظ…ط§ط¯ظٹ + ظƒظˆظ†طھط±ط§ط³طھ + طھظƒط¨ظٹط±/طھطµط؛ظٹط± ظ„ظ…ظ‚ط§ط³ ظ…ظ†ط§ط³ط¨ ظ„ظ„ظ€ OCR
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

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ ط³ظƒط§ظ†ط± ط§ظ„ط¨ط§طھط´: ظ‚طµ ط§ظ„ط¥ظٹطµط§ظ„ ظ…ظ† ط§ظ„طµظˆط±ط© + طھط¹ط¯ظٹظ„ ط§ظ„ظ…ظ†ط¸ظˆط± + طھط­ط³ظٹظ† (ط²ظٹ ظƒط§ظ… ط³ظƒط§ظ†ط±) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// ط§ظ„ظپظƒط±ط©: ط§ظ„ظˆط±ظ‚ط© ط§ظ„ط¨ظٹط¶ط§ (ط§ظ„ط¨ط§طھط´) ط£ظپطھط­ ظ…ظ† ط§ظ„ط®ظ„ظپظٹط© (طھط±ط§ط¨ظٹط²ط©/ط¥ظٹط¯/ط£ط±ط¶ظٹط©). ط¨ظ†ظپطµظ„ ط§ظ„ظˆط±ظ‚ط© ط¨ط­ط¯ظ‘ OtsuطŒ
// ط¨ظ†ط·ظ„ظ‘ط¹ ط£ظƒط¨ط± ظ…ظƒظˆظ‘ظ†طŒ ط¨ظ†ظپط±ط¯ ظ…ظ†ط¸ظˆط±ظ‡ (4 ط£ط±ظƒط§ظ†)طŒ ظˆط¨ط¹ط¯ظٹظ† ط¨ظ†ط´ظٹظ„ ط§ظ„ط¸ظ„ط§ظ„ ظˆظ†ط¨ظٹظ‘ط¶ ط§ظ„ظˆط±ظ‚ط© ظˆظ†ط؛ظ…ظ‘ظ‚ ط§ظ„ظƒطھط§ط¨ط©.
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

// ط®ط· ظ…ط³طھظ‚ظٹظ… q = m*p + k ط¨ط£ظ‚ظ„ ظ…ط±ط¨ط¹ط§طھ ظ…ط¹ ط§ط³طھط¨ط¹ط§ط¯ ط§ظ„ظ†ظ‚ط· ط§ظ„ط´ط§ط±ط¯ط©. null ظ„ظˆ ط§ظ„ظ†ظ‚ط· ظ‚ظ„ظٹظ„ط©.
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

// ط£ط±ظƒط§ظ† ط§ظ„ظˆط±ظ‚ط© [tlx,tly,trx,try,blx,bly,brx,bry] ط¨ط¥ط­ط¯ط§ط«ظٹط§طھ ط§ظ„طµظˆط±ط© ط§ظ„طµط؛ظٹط±ط©
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

  // ط£طµط؛ط± ظ…ط³طھط·ظٹظ„ ظ…ط§ظٹظ„ ظٹط­ظٹط· ط¨ط§ظ„ظˆط±ظ‚ط©
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
  final lp = vertical ? ay : ax; // ط¹ظ„ظ‰ ط·ظˆظ„ ط§ظ„ظˆط±ظ‚ط©
  final lq = vertical ? ax : ay; // ط¹ط±ط¶ ط§ظ„ظˆط±ظ‚ط©
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
  // tl, tr, br, bl
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

// ط¨ظٹط±ط¬ظ‘ط¹ ط£ط±ظƒط§ظ† ظƒظ„ ط¨ط§طھط´ ظپظٹ ط§ظ„طµظˆط±ط© (ط¨ط¥ط­ط¯ط§ط«ظٹط§طھ ط§ظ„طµظˆط±ط© ط§ظ„ط£طµظ„ظٹط©). ظپط§ط¶ظٹ = ظ…ط´ ط¹ط§ط±ظپ ظٹظ„ط§ظ‚ظٹ ظˆط±ظ‚ط©.
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
    // ط§ظ„طµظˆط±ط© ظƒظ„ظ‡ط§ ظˆط±ظ‚ط© (ط³ظƒط±ظٹظ†/ط§ط³ظƒط§ظ†ط± ط¬ط§ظ‡ط²): ط¨ط³ ظ†ط­ط³ظ‘ظ†ظ‡ط§
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
  // ظ„ظˆ ظپظٹ ط£ظƒطھط± ظ…ظ† ط¨ط§طھط´ ظپظٹ ط§ظ„طµظˆط±ط©: ط±طھظ‘ط¨ظ‡ظ… ط´ظ…ط§ظ„ ظ„ظٹظ…ظٹظ† (ط£ظˆ ظپظˆظ‚ ظ„طھط­طھ ظ„ظˆ ظپظˆظ‚ ط¨ط¹ط¶)
  out.sort((a, b) {
    final acx = (a[0] + a[2] + a[4] + a[6]) / 4, bcx = (b[0] + b[2] + b[4] + b[6]) / 4;
    final acy = (a[1] + a[3] + a[5] + a[7]) / 4, bcy = (b[1] + b[3] + b[5] + b[7]) / 4;
    final aw = _dist(a[0], a[1], a[2], a[3]), bw = _dist(b[0], b[1], b[2], b[3]);
    if ((acx - bcx).abs() > math.min(aw, bw) * 0.5) return acx.compareTo(bcx);
    return acy.compareTo(bcy);
  });
  return out;
}

// ط¨ظٹظپط±ط¯ ط§ظ„ظˆط±ظ‚ط© (طھط¹ط¯ظٹظ„ ظ…ظ†ط¸ظˆط±) ظˆظٹط­ط³ظ‘ظ†ظ‡ط§: ظٹط´ظٹظ„ ط§ظ„ط¸ظ„ط§ظ„طŒ ظٹط¨ظٹظ‘ط¶ ط§ظ„ط®ظ„ظپظٹط©طŒ ظٹط؛ظ…ظ‘ظ‚ ط§ظ„ظƒطھط§ط¨ط©
img.Image? _warpAndEnhance(img.Image im, List<double> q) {
  final wq = (_dist(q[0], q[1], q[2], q[3]) + _dist(q[4], q[5], q[6], q[7])) / 2;
  final hq = (_dist(q[0], q[1], q[4], q[5]) + _dist(q[2], q[3], q[6], q[7])) / 2;
  if (wq < 40 || hq < 40) return null;
  final k = math.min(1.0, 2200.0 / math.max(wq, hq));
  final wo = math.max(8, (wq * k).round());
  final ho = math.max(8, (hq * k).round());

  // ظ‚طµ ط§ظ„ظ…ظ†ط·ظ‚ط© ط§ظ„ظ…ط­ظٹط·ط© ط¨ط§ظ„ظˆط±ظ‚ط© ط¨ط³ (ط£ط³ط±ط¹ ظƒطھظٹط±) ظˆطھطµط؛ظٹط±ظ‡ط§ ظ„ظˆ ظƒط¨ظٹط±ط©
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

  // ظپط±ط¯ ط§ظ„ظ…ظ†ط¸ظˆط± (ط§ط³طھظٹظپط§ط، ط«ظ†ط§ط¦ظٹ ط¨ظٹظ† ط§ظ„ط£ط±ظƒط§ظ† ط§ظ„ط£ط±ط¨ط¹ط©)
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

  // طھط­ط³ظٹظ†: ظ†ظ‚ط¯ظ‘ط± ظ„ظˆظ† ط§ظ„ظˆط±ظ‚ط© (ط£ط¹ظ„ظ‰ ظ‚ظٹظ…ط© ظپظٹ ظƒظ„ ظ…ط±ط¨ط¹) ظˆظ†ظ‚ط³ظ… ط¹ظ„ظٹظ‡ = ط´ظٹظ„ ط§ظ„ط¸ظ„ط§ظ„ ظˆطھط¨ظٹظٹط¶ ط§ظ„ط®ظ„ظپظٹط©
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
      // ظ…ظ†ط­ظ†ظ‰ ط£ط؛ظ…ظ‚ ظ„ظ„ظƒطھط§ط¨ط© (gamma) ط¹ظ„ط´ط§ظ† ط§ظ„ط£ط±ظ‚ط§ظ… طھظپط¶ظ„ طھظ‚ظٹظ„ط© ظˆظˆط§ط¶ط­ط©
      final o = math.pow(_cd((n - 30.0) / (240.0 - 30.0), 0.0, 1.0), 1.6).toDouble();
      final c = (o * 255).round();
      res.setPixelRgb(x, y, c, c, c);
    }
  }
  return res;
}

// ط¨طھط´طھط؛ظ„ ظپظٹ Isolate. ط¨طھط±ط¬ظ‘ط¹ ظ…ط³ط§ط±ط§طھ ط§ظ„طµظپط­ط§طھ (طµظپط­ط© ظ„ظƒظ„ ط¨ط§طھط´). ظپط§ط¶ظٹ = ط§ط³طھط®ط¯ظ… ط§ظ„طµظˆط±ط© ط§ظ„ط£طµظ„ظٹط©.
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


// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ ظ…ظ„ط، ط§ظ„ظ‚ط§ظ„ط¨ ظ…ط¨ط§ط´ط±ط© ط¬ظˆظ‡ ظ…ظ„ظپ ط§ظ„ظ€ xlsx (ظ…ظ† ط؛ظٹط± ظ…ط§ ظ†ط¨ظˆظ‘ط¸ ط§ظ„طھظ†ط³ظٹظ‚) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
Future<List<int>?> loadTemplateBytes() async {
  try {
    final mine = File('$baseDir/template.xlsx');
    if (mine.existsSync()) return mine.readAsBytesSync();
    final exact = File('$baseDir/ظ†طھظٹط¬ط© 00.xlsx');
    if (exact.existsSync()) return exact.readAsBytesSync();
    final d = Directory(baseDir);
    if (d.existsSync()) {
      for (final e in d.listSync().whereType<File>()) {
        final n = baseName(e.path);
        if (n.toLowerCase().endsWith('.xlsx') &&
            n.contains('00') &&
            !n.startsWith('~') &&
            !n.startsWith('.')) {
          return e.readAsBytesSync();
        }
      }
    }
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

double _readNum(String xml, String col, int row) {
  final rm = _findRow(xml, row);
  if (rm == null) return 0.0;
  for (final m in _cellRe.allMatches(rm.group(0)!)) {
    final cell = m.group(0)!;
    final open = RegExp(r'<c\b[^>]*?>').firstMatch(cell)!.group(0)!;
    final r = RegExp(r'\br="([A-Z]+)\d+"').firstMatch(open);
    if (r == null || r.group(1) != col) continue;
    final t = RegExp(r'\bt="([^"]+)"').firstMatch(open)?.group(1);
    if (t != null && t != 'n') return 0.0;
    final v = RegExp(r'<v>([^<]*)</v>').firstMatch(cell)?.group(1);
    return double.tryParse(v ?? '') ?? 0.0;
  }
  return 0.0;
}

// ط¨ظٹظƒطھط¨ ط®ظ„ظٹط© ظˆط§ط­ط¯ط© ظ…ط­ط§ظپط¸ ط¹ظ„ظ‰ ط§ظ„ظ€ style ط¨طھط§ط¹ظ‡ط§. build(style, oldFormulaElement) -> xml ط§ظ„ط®ظ„ظٹط©
String _setCell(String xml, String col, int row,
    String Function(String ref, String? style, String? formula) build) {
  final rm = _findRow(xml, row);
  if (rm == null) throw Exception('row $row missing');
  var rowXml = rm.group(0)!;
  final ref = '$col$row';
  final target = _colNum(col);
  RegExpMatch? existing, before;
  for (final m in _cellRe.allMatches(rowXml)) {
    final cell = m.group(0)!;
    final open = RegExp(r'<c\b[^>]*?>').firstMatch(cell)!.group(0)!;
    final r = RegExp(r'\br="([A-Z]+)\d+"').firstMatch(open);
    if (r == null) continue;
    final cn = _colNum(r.group(1)!);
    if (cn == target) {
      existing = m;
      break;
    }
    if (cn > target && before == null) before = m;
  }
  String? style, formula;
  if (existing != null) {
    final cell = existing.group(0)!;
    final open = RegExp(r'<c\b[^>]*?>').firstMatch(cell)!.group(0)!;
    style = RegExp(r'\bs="(\d+)"').firstMatch(open)?.group(1);
    formula = RegExp(r'<f\b[^>]*?(?:/>|>.*?</f>)', dotAll: true).firstMatch(cell)?.group(0);
  }
  final newCell = build(ref, style, formula);
  if (existing != null) {
    rowXml = rowXml.replaceRange(existing.start, existing.end, newCell);
  } else if (before != null) {
    rowXml = rowXml.replaceRange(before.start, before.start, newCell);
  } else if (rowXml.endsWith('/>')) {
    rowXml = '${rowXml.substring(0, rowXml.length - 2)}>$newCell</row>';
  } else {
    final i = rowXml.lastIndexOf('</row>');
    rowXml = rowXml.replaceRange(i, i, newCell);
  }
  return xml.replaceRange(rm.start, rm.end, rowXml);
}

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

// ط¨ظٹط¨ظ†ظٹ طµظپط­ط© ط§ظ„ظ†طھط§ظٹط¬: طµظپ ط§ظ„ط¹ظ†ط§ظˆظٹظ† ط²ظٹ ظ…ط§ ظ‡ظˆطŒ ظˆطµظپ ظ„ظƒظ„ ظˆط­ط¯ط© ط¨ظ†ظپط³ طھظ†ط³ظٹظ‚ ط£ظˆظ„ طµظپ ط¨ظٹط§ظ†ط§طھ ظپظٹ ط§ظ„ظ‚ط§ظ„ط¨
// (ط¹ط¯ط¯ ط§ظ„طµظپظˆظپ = ط¹ط¯ط¯ ط§ظ„ظˆط­ط¯ط§طھ ط§ظ„ظ„ظٹ ط§طھط¹ط±ظپطھ ظ…ظ† ط§ظ„ظپط­طµطŒ ط¨ظٹط²ظٹط¯ ط£ظˆ ط¨ظٹظ‚ظ„ ط­ط³ط¨ ط§ظ„ط¹ط¯ط¯)
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
    if (f.name == 'xl/calcChain.xml') continue; // ط¨ظٹطھط¹ظ…ظ„ ظ…ظ† ط¬ط¯ظٹط¯طŒ ظˆظƒط§ظ† ظ‡ظٹط¨ظˆظ‘ط¸ ط§ظ„ظ…ظ„ظپ ط¨ط¹ط¯ طھط؛ظٹظٹط± ط§ظ„طµظپظˆظپ
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

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ ظ‚ظ†ظˆط§طھ ط§ظ„ظ†ط¸ط§ظ… (ظˆط§طھط³ط§ط¨ ظ…ط¨ط§ط´ط± / ظپطھط­ ظپظˆظ„ط¯ط±) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const MethodChannel _chan = MethodChannel('pos/native');

String todayToken() {
  final n = DateTime.now();
  return '${n.day}.${n.month}';
}

String _normAr(String s) => s
    .replaceAll(RegExp('[ط£ط¥ط¢]'), 'ط§')
    .replaceAll('ط©', 'ظ‡')
    .replaceAll('ظ‰', 'ظٹ')
    .toLowerCase();

// ظ…ظ„ظپ PDF طھظ‚ظپظٹظ„ ط§ط³ظ…ظ‡ ظپظٹظ‡ طھط§ط±ظٹط® ط§ظ„ظ†ظ‡ط§ط±ط¯ظ‡ (ط¨ظٹظ‚ط¨ظ„ 5.11 / 05.11 / 5-11 / 5_11)
bool isTodayAlexPdf(String path) {
  final name = _normAr(baseName(path));
  if (!name.endsWith('.pdf')) return false;
  if (!(name.contains('ط§ط³ظƒظ†ط¯ط±ظٹ') || name.contains('ظ†طھظٹط¬') || name.contains('طھظ‚ظپظٹظ„') || name.contains('ط§ظ‚ظپط§ظ„'))) return false;
  final n = DateTime.now();
  final re = RegExp('(?<!\\d)0?${n.day}[._\\- ]0?${n.month}(?!\\d)');
  return re.hasMatch(name);
}

List<File> findTodayAlexPdfs(List<String> dirs) {
  final out = <File>[];
  final seen = <String>{};
  for (final d in dirs) {
    if (d == outputDir) continue; // ظ…ط§ ظ†ظ‚ط±ط§ط´ ط§ظ„ظ…ظ„ظپط§طھ ط§ظ„ظ„ظٹ ط§ظ„طھط·ط¨ظٹظ‚ ظ†ظپط³ظ‡ ط·ظ„ظ‘ط¹ظ‡ط§
    try {
      final dir = Directory(d);
      if (!dir.existsSync()) continue;
      for (final e in dir.listSync().whereType<File>()) {
        if (isTodayAlexPdf(e.path) && seen.add(e.path)) out.add(e);
      }
    } catch (_) {}
  }
  out.sort((a, b) => b.lastModifiedSync().compareTo(a.lastModifiedSync()));
  return out;
}

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ ظ†ظ…ظˆط°ط¬ ط§ظ„ط¥ظٹطµط§ظ„ â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
class ReceiptItem {
  String path;
  String? tid;
  int? orderNum;
  int? row;
  String? unitName;
  double amount;
  String? date;
  String key; // ط§ط³ظ… ط§ظ„ظ…ظ„ظپ ط§ظ„ط£طµظ„ظٹ + ط­ط¬ظ…ظ‡ (ظ„ط­ظپط¸ ط§ظ„طھطµط­ظٹط­ط§طھ ط§ظ„ظٹط¯ظˆظٹط©)
  bool ok; // ط§ظ„ظ…ط¨ظ„ط؛ ظ…ط¤ظƒط¯طں
  String how; // terminal | fuzzy | name/merchant | auto-number | conflict | extra-batch | manual
  List<double> batches;
  int? suffix;
  String? origTid;
  bool conflict;
  bool skip;
  double? alt; // ظ‚ط±ط§ط،ط© طھط§ظ†ظٹط© ظ…ط®طھظ„ظپط© ظ„ظ„ظ…ط¨ظ„ط؛ (ظ„ظ„ظ…ط±ط§ط¬ط¹ط©)
  bool dupSuspect; // ظ…ط´ظƒظˆظƒ ط¥ظ†ظ‡ ظ…ظƒط±ط±: ط¨ظٹط¯ط®ظ„ ط§ظ„ظ…ط±ط§ط¬ط¹ط© ظˆظ…ط´ ط¨ظٹطھط­ط³ط¨ ظ„ط­ط¯ ظ…ط§ طھظ‚ط±ط±
  bool userExcluded; // ط§ظ†طھ ط§ط³طھط¨ط¹ط¯طھظ‡ ظپظٹ ط§ظ„ظ…ط±ط§ط¬ط¹ط©
  String? dupDecision; // 'keep' | 'drop' (ظ‚ط±ط§ط±ظƒ ط¨ظٹطھط­ظپط¸ ظ„ظ„ظ…ط±ط© ط§ظ„ط¬ط§ظٹط©)
  int? dupOf; // ط±ظ‚ظ… ط§ظ„ظ…ط§ظƒظٹظ†ط© ط§ظ„ظ„ظٹ ط´ط¨ظ‡ظ‡
  String? unusual; // ط¥ط¬ظ…ط§ظ„ظٹ ط¨ط¹ظٹط¯ ط¹ظ† ط§ظ„ظ…ط¹طھط§ط¯ (ظ„ظ„ظ…ط±ط§ط¬ط¹ط©)
  bool ignoredDoc; // طµظˆط±ط© ط§طھط¬ط§ظ‡ظ„ظ‡ط§ ط§ظ„طھط·ط¨ظٹظ‚ (ط´ظƒظ„ظ‡ط§ ط¥ظٹطµط§ظ„ ط¯ظپط¹/ظ…ط³طھظ†ط¯): ط¨طھط¯ط®ظ„ ط§ظ„ظ…ط±ط§ط¬ط¹ط©
  String? ign; // 'keep' | 'drop' ظ‚ط±ط§ط±ظƒ ظپظٹ ط§ظ„طµظˆط±ط© ط§ظ„ظ…طھط¬ط§ظ‡ظ„ط©
  bool dateOdd; // طھط§ط±ظٹط®ظ‡ ظ…ط®طھظ„ظپ ط¹ظ† ط£ط؛ظ„ط¨ ط§ظ„ط¥ظٹطµط§ظ„ط§طھ
  String? dateMain; // ط§ظ„طھط§ط±ظٹط® ط§ظ„ط£ط؛ظ„ط¨
  String? dateFix; // ط§ظ„طھط§ط±ظٹط® ط§ظ„ظ„ظٹ ط§ظ†طھ ط£ظƒط¯طھظ‡
  int? txn; // ط¹ط¯ط¯ ظ…ط¹ط§ظ…ظ„ط§طھ ط§ظ„ط¨ط§طھط´ ط§ظ„ظ…ظ‚ط±ظˆط، ظ…ظ† ط¥ظٹطµط§ظ„ ط§ظ„ط¥ظ‚ظپط§ظ„

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
  }) : batches = batches ?? <double>[];

  // ط¯ط§ط®ظ„ ظپظٹ ط§ظ„ط¥ط¬ظ…ط§ظ„ظٹط§طھ ظˆط§ظ„ظ€ PDFطں
  bool get counts => !skip && !dupSuspect && !userExcluded && !ignoredDoc;

  // ط§ظ„ظ…ط§ظƒظٹظ†ط©/ط§ظ„ظˆط­ط¯ط© ظ…طھط­ط¯ط¯ط© ط¨ط·ط±ظٹظ‚ط© ظ…ط´ ظ…ط¤ظƒط¯ط© (طھظ„ظ‚ط§ط¦ظٹ / ط¨ط§ظ„ط§ط³ظ… ظپظ‚ط·)
  bool get weakId =>
      !(how == 'terminal' || how == 'fuzzy' || how == 'manual' || how == 'extra-batch');

  bool get needsReview =>
      ignoredDoc ||
      dupSuspect ||
      dateOdd ||
      (!userExcluded &&
          (amount <= 0 ||
              row == null ||
              !ok ||
              alt != null ||
              unusual != null ||
              weakId));
}

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ ط§ظ„طھطµط­ظٹط­ط§طھ ط§ظ„ظٹط¯ظˆظٹط© (ط¨طھظپط¶ظ„ ظ…ط­ظپظˆط¸ط©) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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
    };
    File(fixPath)
      ..createSync(recursive: true)
      ..writeAsStringSync(jsonEncode(m));
  } catch (_) {}
}

// طµظˆط±ط© ظƒط§ظ† ط§ظ„طھط·ط¨ظٹظ‚ ظ…طھط¬ط§ظ‡ظ„ظ‡ط§: ط§ط¯ط®ظ„ظ‡ط§ ظƒط¨ط§طھط´ (ظ‡طھظƒظ…ظ‘ظ„ ط§ظ„ظ…ط¨ظ„ط؛ ظˆط§ظ„ظ…ط§ظƒظٹظ†ط© ط¨ط¥ظٹط¯ظƒ)
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

// ط§ط­طھط³ط§ط¨ ط¥ظٹطµط§ظ„ (ظ…ط´ظƒظˆظƒ ط£ظˆ ظ…ط³طھط¨ط¹ط¯) ظƒط¨ط§طھط´ ط¥ط¶ط§ظپظٹ ظ„ظ„ظ…ط§ظƒظٹظ†ط©
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

// ط¥ط¬ظ…ط§ظ„ظٹ ظ…ط§ظƒظٹظ†ط© ط¨ط¹ظٹط¯ ط¬ط¯ط§ظ‹ ط¹ظ† ظ…طھظˆط³ط· ط§ظ„ط£ظٹط§ظ… ط§ظ„ظ„ظٹ ظپط§طھطھ = ط؛ط§ظ„ط¨ط§ظ‹ ظ‚ط±ط§ط،ط© ط؛ظ„ط· -> ظ…ط±ط§ط¬ط¹ط©
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
        final msg = 'ط¥ط¬ظ…ط§ظ„ظٹ ظ…ط§ظƒظٹظ†ط© ${e.key} (${e.value.toStringAsFixed(2)}) ط¨ط¹ظٹط¯ ط¹ظ† ظ…طھظˆط³ط· '
            'ط§ظ„ط£ظٹط§ظ… ط§ظ„ظ„ظٹ ظپط§طھطھ (${avg.toStringAsFixed(0)}) - ط§طھط£ظƒط¯ ظ…ظ† ط§ظ„ظ…ط¨ظ„ط؛';
        for (final r in items) {
          if (r.orderNum == e.key && r.how != 'manual') r.unusual = msg;
        }
      }
    }
    if (reportDate != 'ط§ظ„ظٹظˆظ…') {
      hist[reportDate] = {for (final e in today.entries) '${e.key}': e.value};
      while (hist.length > 30) {
        hist.remove(hist.keys.first);
      }
      cfg['hist'] = hist;
      _writeCfg(cfg);
    }
  } catch (_) {}
}


// â”€â”€â”€â”€â”€â”€â”€â”€â”€ ط§ظ„ط¹ط¯ ط§ظ„ط³ط±ظٹط¹: طµظˆط± ط§ظ„طµظپط­ط§طھ ط§ظ„ط£طµظ„ظٹط© (JPEG) ظ…ظ† ط§ظ„ظ€ PDF ط§ظ„ظ…ظ…ط³ظˆط­ + ط¹ط¯ظ‘ ط¨ط§ظ„ظ‡ظٹظƒظ„ â”€â”€â”€â”€â”€â”€â”€â”€â”€
bool _loadFast() {
  try {
    return _readCfg()['fastCount'] != false;
  } catch (_) {
    return true;
  }
}

bool fastCount = _loadFast();

void toggleFastCount() {
  fastCount = !fastCount;
  try {
    final m = _readCfg();
    m['fastCount'] = fastCount;
    _writeCfg(m);
  } catch (_) {}
}

// ط¨ظٹط·ظ„ظ‘ط¹ طµظˆط± ط§ظ„طµظپط­ط§طھ (JPEG) ط¬ط§ظ‡ط²ط© ظ…ظ† PDF ظ…ظ…ط³ظˆط­ (CamScanner ظˆط؛ظٹط±ظ‡): ظ…ظ† ط؛ظٹط± raster ظˆظ„ط§ PNG
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
  double h = 3000;
  final Map<String, int> types = {};
}

// ط¹ط¯ظ‘ ط§ظ„ط¥ظٹطµط§ظ„ط§طھ ط¨ط§ظ„ظ‡ظٹظƒظ„: ظƒظ„ ط¥ظٹطµط§ظ„ ظ„ظ‡ ط´ط¹ط§ط± + Terminal ID + RRN + AMOUNT + Approved (ظ…ط±ط© ظˆط§ط­ط¯ط©)
// ط¨ظ†ط¬ظ…ط¹ ط§ظ„ط¹ظ„ط§ظ…ط§طھ ط¯ظٹ ظˆظ†ظ‚ط³ظ…ظ‡ط§ ط£ط¹ظ…ط¯ط©طŒ ظˆظپظٹ ظƒظ„ ط¹ظ…ظˆط¯ ط£ظƒطھط± ط¹ظ„ط§ظ…ط© ظ…طھظƒط±ط±ط© = ط¹ط¯ط¯ ط§ظ„ط¥ظٹطµط§ظ„ط§طھ
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
    if (amtRe.hasMatch(u) && !amtBad.hasMatch(u)) add('amount');
    if (u.contains('APPROVED')) add('appr');
  }
  // ط¥ط²ط§ظ„ط© ط§ظ„طھظƒط±ط§ط± (ظ†ظپط³ ط§ظ„ط¹ظ„ط§ظ…ط© ط§طھظ‚ط±طھ ظ…ط±طھظٹظ†)
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
  // ط£ط¹ظ…ط¯ط©
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
  // ظ„ظƒظ„ ط¹ظ…ظˆط¯: ظƒظ„ ط¹ظ„ط§ظ…ط© ط¨طھط¸ظ‡ط± ظ…ط±ط© ظˆط§ط­ط¯ط© ظپظٹ ظƒظ„ ط¥ظٹطµط§ظ„طŒ ظپط£ظƒط¨ط± ط¹ط¯ط¯ ظ„ط£ظٹ ط¹ظ„ط§ظ…ط© = ط¹ط¯ط¯ ط§ظ„ط¥ظٹطµط§ظ„ط§طھ ظپظٹ ط§ظ„ط¹ظ…ظˆط¯
  // (ظ„ظˆ ط¹ظ„ط§ظ…ط© ط§طھظ‚ط±طھ ظ†ط§ظ‚طµط© ط§ظ„طھط§ظ†ظٹط© ط¨طھط¹ظˆظ‘ط¶طŒ ظˆظ…ط´ ط¨ظٹطھط£ط«ط± ط¨ظ…ظٹظ„ ط§ظ„ط¥ظٹطµط§ظ„ ظˆظ„ط§ ط¨طھط±طھظٹط¨ ط§ظ„ظ‚ط±ط§ط،ط©)
  var n = 0;
  for (final c in cols) {
    final cnt = <String, int>{};
    for (final a in c) {
      cnt[a.t] = (cnt[a.t] ?? 0) + 1;
    }
    n += cnt.values.reduce((a, b) => a > b ? a : b);
  }
  out.sm = n;
  return out;
}

bool sameDay(ReceiptItem a, ReceiptItem b) =>
    a.date == null || b.date == null || a.date == b.date;

// ط¨ظٹط¬ظٹط¨ ظƒظ„ ط§ظ„طھط·ط¨ظٹظ‚ط§طھ ط§ظ„ظ„ظٹ طھظ‚ط¯ط± طھظپطھط­ ظپظˆظ„ط¯ط± (ظ…ط¯ظٹط± ظ…ظ„ظپط§طھ) ظˆظٹط®ظ„ظ‘ظٹظƒ طھط®طھط§ط± ط£ظٹ ظˆط§ط­ط¯طŒ ظ…ط´ ط§ظ„ط§ظپطھط±ط§ط¶ظٹ ط¨ط³
Future<bool> chooseAndOpenFolder(BuildContext context, String path) async {
  var handlers = <Map<String, String>>[];
  try {
    final r = await _chan.invokeMethod<List<dynamic>>('folderHandlers', {'path': path});
    if (r != null) {
      handlers = [
        for (final e in r)
          {for (final en in (e as Map).entries) en.key.toString(): en.value.toString()}
      ];
    }
  } catch (_) {}
  // ظƒظ„ طھط·ط¨ظٹظ‚ ظ…ط±ط© ظˆط§ط­ط¯ط© (ط£ظˆظ„ ط·ط±ظٹظ‚ط© ط§طھط¬ط±ط¨طھ ظ‡ظٹ ط§ظ„ظ„ظٹ ط´ط؛ط§ظ„ط©)طŒ ظˆط¨ظ†ط´ظٹظ„ Termux ظ„ط£ظ†ظ‡ ظ…ط´ ط¨ظٹط¹ط±ط¶ ظپظˆظ„ط¯ط±ط§طھ
  final seenPkgs = <String>{};
  handlers = handlers
      .where((h) => h['pkg'] != 'com.termux' && seenPkgs.add(h['pkg'] ?? ''))
      .toList();
  if (!context.mounted) return false;
  final picked = await showModalBottomSheet<Map<String, String>>(
    context: context,
    isScrollControlled: true,
    builder: (c) => SafeArea(
      child: ListView(
        shrinkWrap: true,
        children: [
          const ListTile(
            title: Text('ط§ظپطھط­ ط§ظ„ظپظˆظ„ط¯ط± ط¨ظˆط§ط³ط·ط©', style: TextStyle(fontWeight: FontWeight.bold)),
          ),
          for (final h in handlers)
            ListTile(
              leading: h['icon'] != null && h['icon']!.isNotEmpty ? Image.memory(base64Decode(h['icon']!), width: 32, height: 32) : const Icon(Icons.folder, size: 32),
              title: Text(h['label'] ?? h['pkg'] ?? '', style: const TextStyle(fontWeight: FontWeight.bold)),
              onTap: () => Navigator.pop(c, h),
            ),
          ListTile(
            leading: const Icon(Icons.folder_special, size: 32),
            title: const Text('ظ…طھطµظپط­ ظ…ظ„ظپط§طھ ط§ظ„ظ†ط¸ط§ظ…', style: const TextStyle(fontWeight: FontWeight.bold)),
            onTap: () => Navigator.pop(c, <String, String>{'pkg': '', 'cls': '', 'kind': 'tree'}),
          ),
        ],
      ),
    ),
  );
  if (picked == null) return true; // ط§ظ„ظ…ط³طھط®ط¯ظ… ظ‚ظپظ„ ط§ظ„ظ‚ط§ظٹظ…ط©
  try {
    final ok = await _chan.invokeMethod<bool>('openFolderWith', {
      'path': path,
      'pkg': picked['pkg'] ?? '',
      'cls': picked['cls'] ?? '',
      'kind': picked['kind'] ?? '',
    });
    return ok == true;
  } catch (_) {
    return false;
  }
}

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ ط§ظ„ط´ط§ط´ط© ط§ظ„ط±ط¦ظٹط³ظٹط© â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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
  String status = 'ط§ط¶ط؛ط· "ط¨ط¯ط، ط§ظ„ظ…ط¹ط§ظ„ط¬ط© ط§ظ„ط´ط§ظ…ظ„ط©" ظ„ط³ط­ط¨ ط§ظ„طµظˆط± ظˆطھظˆظ„ظٹط¯ ط§ظ„ط¥ظƒط³ظٹظ„ ظˆط§ظ„ظ€ PDF.';
  final List<ReceiptItem> processedItems = [];
  String? standardXlsxPath;
  String? generatedPdfPath;
  List<String> extraPdfPaths = [];
  List<int> missingTerminals = [];
  final List<String> warnings = [];
  final List<String> excludedPay = [];
  bool autoRotate = false; // ط£ط¨ط·ط£: ط¨ظٹط¬ط±ظ‘ط¨ طھط¯ظˆظٹط± ظƒظ„ طµظˆط±ط© ظ…ط´ ط¥ظٹطµط§ظ„
  String sourceNote = '';
  bool reviewOnly = false;

  @override
  void initState() {
    super.initState();
    WidgetsBinding.instance.addObserver(this);
    days = sharedDays.value;
    sharedDays.addListener(_syncDays);
    wipeSignal.addListener(_onWipe);
    customDir = loadFolder();
    customOut = loadOutDir();
    if (customOut != null && customOut!.isNotEmpty) outputDir = customOut!;
    // ط£ظˆظ„ ظ…ط§ ط§ظ„طھط·ط¨ظٹظ‚ ظٹظپطھط­: ط§ط·ظ„ط¨ ط¥ط°ظ† ط§ظ„ظ…ظ„ظپط§طھ (ظ„ظˆ ظ…ط´ ظ…ظ…ظ†ظˆط­) ظˆط§ظ†ط²ظ‘ظ„ ط´ظٹطھ ط§ظ„ط¥ط¹ط¯ط§ط¯ط§طھ ظˆط´ظٹطھ ط§ظ„ظ†طھظٹط¬ط© طھظ„ظ‚ط§ط¦ظٹ
    WidgetsBinding.instance.addPostFrameCallback((_) async {
      try {
        if (!await Permission.manageExternalStorage.isGranted) {
          await Permission.manageExternalStorage.request();
        }
      } catch (_) {}
      await prepareDefaultFiles();
    });
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
      status = 'طھظ… ظ…ط³ط­ ظƒظ„ ط­ط§ط¬ط©. ط§ط¶ط؛ط· "ط¨ط¯ط، ط§ظ„ظ…ط¹ط§ظ„ط¬ط© ط§ظ„ط´ط§ظ…ظ„ط©" ظ„ط¨ط¯ط، ظٹظˆظ… ط¬ط¯ظٹط¯.';
    });
  }

  Future<void> _onSwipe(ReceiptItem it, bool approve) async {
    if (busy) return;
    if (approve) {
      if (it.dupSuspect) {
        keepAsExtraBatch(it, processedItems);
        await rebuildFiles();
      } else if (it.dateOdd) {
        it.dateFix = it.date;
        it.dateOdd = false;
        saveFix(it);
        await rebuildFiles();
      } else {
        if (it.amount <= 0 || it.row == null) {
          ScaffoldMessenger.of(context).showSnackBar(SnackBar(
              content: w.Text(_l(context, 'ظƒظ…ظ‘ظ„ ط§ظ„ظ…ط¨ظ„ط؛ ظˆط§ظ„ظ…ط§ظƒظٹظ†ط© ظ…ظ† ط²ط± ط§ظ„ظ‚ظ„ظ… ط§ظ„ط£ظˆظ„.',
                  'Fill in the amount and machine with the pencil button first.'))));
          return;
        }
        it.ok = true;
        it.alt = null;
        it.unusual = null;
        if (it.how != 'extra-batch') it.how = 'manual';
        saveFix(it);
        setState(() {});
      }
    } else {
      if (it.dupSuspect) {
        it.dupSuspect = false;
        it.userExcluded = true;
        it.dupDecision = 'drop';
        saveFix(it);
        await rebuildFiles();
      } else if (it.dateOdd) {
        if (it.dateMain != null) it.date = it.dateMain;
        it.dateFix = it.date;
        it.dateOdd = false;
        saveFix(it);
        await rebuildFiles();
      } else {
        await editReceipt(it);
      }
    }
  }

  // ط³ط­ط¨ ط¥ظٹطµط§ظ„ ط§ظ„ظ…ط±ط§ط¬ط¹ط©: طµط­ âœ“ / ط؛ظ„ط· âœ— (ط؛ظ„ط· = ط§ط³طھط¨ط¹ط§ط¯ ظ„ظˆ ظ…ظƒط±ط±طŒ ظˆط¥ظ„ط§ ط¨ظٹظپطھط­ ط§ظ„طھطµط­ظٹط­)
  Widget _swipeWrap(ReceiptItem item, Widget card) {
    if (item.ignoredDoc || item.userExcluded || !item.needsReview) return card;
    final okLabel = item.dupSuspect
        ? _l(context, 'ط§ط­طھط³ط¨ظ‡ ط¨ط§طھط´ âœ“', 'Count it âœ“')
        : (item.dateOdd ? _l(context, 'ط³ظٹط¨ طھط§ط±ظٹط®ظ‡ âœ“', 'Keep date âœ“') : _l(context, 'طµط­ âœ“', 'OK âœ“'));
    final badLabel = item.dupSuspect
        ? _l(context, 'ظ…ظƒط±ط± - ط§ط³طھط¨ط¹ط¯ظ‡ âœ—', 'Duplicate âœ—')
        : (item.dateOdd ? _l(context, 'طھط§ط±ظٹط® ط§ظ„ط£ط؛ظ„ط¨ âœ—', 'Use main date âœ—') : _l(context, 'ط؛ظ„ط· - طµط­ظ‘ط­ظ‡ âœ—', 'Wrong - fix âœ—'));
    return Dismissible(
      key: ValueKey('sw_${item.path}'),
      direction: busy ? DismissDirection.none : DismissDirection.horizontal,
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
    sharedDays.removeListener(_syncDays);
    wipeSignal.removeListener(_onWipe);
    WidgetsBinding.instance.removeObserver(this);
    super.dispose();
  }

  Future<void> pickFolder() async {
    final perm = await Permission.manageExternalStorage.request();
    if (!perm.isGranted) {
      setState(() => status = 'ظ„ط§ط²ظ… طھط³ظ…ط­ ط¨ط§ظ„ظˆطµظˆظ„ ظ„ظƒظ„ ط§ظ„ظ…ظ„ظپط§طھ ط§ظ„ط£ظˆظ„.');
      return;
    }
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
    if (!perm.isGranted) {
      setState(() => status = 'ظ„ط§ط²ظ… طھط³ظ…ط­ ط¨ط§ظ„ظˆطµظˆظ„ ظ„ظƒظ„ ط§ظ„ظ…ظ„ظپط§طھ ط§ظ„ط£ظˆظ„.');
      return;
    }
    if (!mounted) return;
    final res = await Navigator.push<String>(
      context,
      MaterialPageRoute(builder: (_) => FolderPicker(start: customOut ?? root)),
    );
    if (res != null) {
      try {
        final t = File('$res/.write_test')..writeAsStringSync('x');
        t.deleteSync();
      } catch (_) {
        setState(() => status = 'ظ…ط´ ظ‚ط§ط¯ط± ط£ظƒطھط¨ ظپظٹ ط§ظ„ظپظˆظ„ط¯ط± ط¯ظ‡: $res');
        return;
      }
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
        _runPipelineInner(manualPaths: paths);
      }
    } catch (e) {
      _dbg('pickManualImages error: $e');
    }
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
      SnackBar(content: Text('ظ…ظ‚ط¯ط±طھط´ ط£ظپطھط­ ط§ظ„ظپظˆظ„ط¯ط±. ط§ظ„ظ…ط³ط§ط± ط§طھظ†ط³ط®: $path')),
    );
  }

  // ط¥ط±ط³ط§ظ„ ط§ظ„ط¥ظƒط³ظٹظ„ ظˆط§ظ„ظ€ PDF (ط§ظ„ظ…ط³طھط®ط¯ظ… ط¨ظٹط®طھط§ط± ظˆط§طھط³ط§ط¨ ظ…ظ† ظ‚ط§ظٹظ…ط© ط§ظ„ظ…ط´ط§ط±ظƒط©)
  Future<void> sendToWhatsApp() async {
    final x = latestOutput('.xlsx');
    final p = latestOutput('.pdf');
    final files = <XFile>[if (x != null) XFile(x), if (p != null) XFile(p)];
    if (files.isEmpty) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('ظ„ظ… ظٹطھظ… ط¥ظ†ط´ط§ط، ظ…ظ„ظپط§طھ ط¨ط¹ط¯ ظ„ظ…ط´ط§ط±ظƒطھظ‡ط§.')),
      );
      return;
    }
    if (!await confirmSend(context,
        review: processedItems.where((e) => e.needsReview).length, missing: missingTerminals.length)) {
      return;
    }
    await Share.shareXFiles(files, text: 'طھظ‚ط±ظٹط± ط§ظ„ط¥ظٹطµط§ظ„ط§طھ');
  }

  // ط£ط­ط¯ط« ظ…ظ„ظپ ط¥ط®ط±ط§ط¬ (ط¥ظƒط³ظٹظ„/PDF): ظ…ظ† ط¢ط®ط± ظ…ط¹ط§ظ„ط¬ط©طŒ ط£ظˆ ط¢ط®ط± ظ…ظ„ظپ "ظ†طھظٹط¬ط© ..." ظپظٹ ظپظˆظ„ط¯ط± ط§ظ„ط¥ط®ط±ط§ط¬
  String? latestOutput(String ext) {
    final cur = ext == '.xlsx' ? standardXlsxPath : generatedPdfPath;
    if (cur != null && File(cur).existsSync()) return cur;
    try {
      final d = Directory(outputDir);
      if (!d.existsSync()) return null;
      final fs = d
          .listSync()
          .whereType<File>()
          .where((f) {
            final n = _normAr(baseName(f.path));
            return n.endsWith(ext) && _isOutputFileName(baseName(f.path));
          })
          .toList()
        ..sort((a, b) => b.lastModifiedSync().compareTo(a.lastModifiedSync()));
      return fs.isEmpty ? null : fs.first.path;
    } catch (_) {
      return null;
    }
  }

  Future<void> openFile(String? path, String mime) async {
    if (path == null || !File(path).existsSync()) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('ط§ظ„ظ…ظ„ظپ ظ…ط´ ظ…ظˆط¬ظˆط¯. ط´ط؛ظ‘ظ„ ط§ظ„ظ…ط¹ط§ظ„ط¬ط© ط§ظ„ط£ظˆظ„.')),
      );
      return;
    }
    try {
      final ok = await _chan.invokeMethod<bool>('openFile', {'path': path, 'mime': mime});
      if (ok == true) return;
    } catch (_) {}
    if (!mounted) return;
    ScaffoldMessenger.of(context).showSnackBar(
      const SnackBar(content: Text('ظ…ظپظٹط´ ط¨ط±ظ†ط§ظ…ط¬ ظٹظپطھط­ ط§ظ„ظ…ظ„ظپ ط¯ظ‡.')),
    );
  }

  bool _isAppOutput(String path) {
    final n = baseName(path);
    final l = n.toLowerCase();
    return n == 'review_report.txt' ||
        n == 'ocr_debug.txt' ||
        _isOutputFileName(n);
  }

  Future<bool> _confirm(String title, String body) async {
    final r = await showDialog<bool>(
      context: context,
      builder: (ctx) => AlertDialog(
        title: Text(title),
        content: Text(body),
        actions: [
          TextButton(onPressed: () => Navigator.pop(ctx, false), child: const Text('ط¥ظ„ط؛ط§ط،')),
          FilledButton(onPressed: () => Navigator.pop(ctx, true), child: const Text('ظ…ط³ط­')),
        ],
      ),
    );
    return r == true;
  }

  // ظ…ط³ط­ ط§ظ„ظƒظ„: ظ…ظ„ظپط§طھ ظپظˆظ„ط¯ط± ط§ظ„ط¥ط®ط±ط§ط¬ + ط§ظ„طµظˆط± ط§ظ„ظ…ط±طھط¨ط©
  Future<void> clearAll() async {
    final protectedDirs = <String>{
      root,
      baseDir,
      '$root/Download',
      '$root/Documents',
      '$root/DCIM',
      '$root/Pictures',
    };
    final isProtected = protectedDirs.contains(outputDir);
    final outFiles = <File>[];
    try {
      final d = Directory(outputDir);
      if (d.existsSync()) {
        for (final e in d.listSync().whereType<File>()) {
          if (!isProtected || _isAppOutput(e.path)) outFiles.add(e);
        }
      }
    } catch (_) {}
    final rawFiles = <File>[];
    try {
      final d = Directory(rawImagesDir);
      if (d.existsSync()) rawFiles.addAll(d.listSync().whereType<File>());
    } catch (_) {}
    if (outFiles.isEmpty && rawFiles.isEmpty) {
      setState(() => status = 'ظ…ظپظٹط´ ط­ط§ط¬ط© ظ„ظ„ظ…ط³ط­.');
      return;
    }
    final ok = await _confirm(
      'ظ…ط³ط­ ط§ظ„ظƒظ„',
      'ظ‡طھطھظ…ط³ط­ ${outFiles.length} ظ…ظ„ظپ ظ…ظ† ط§ظ„ط¥ط®ط±ط§ط¬:\n$outputDir\n\n'
          'ظˆ ${rawFiles.length} طµظˆط±ط© ظ…ظ† ط§ظ„طµظˆط± ط§ظ„ظ…ط±طھط¨ط©.'
          '${isProtected ? '\n\n(ظپظˆظ„ط¯ط± ط§ظ„ط¥ط®ط±ط§ط¬ ط¯ظ‡ ط¹ط§ظ…طŒ ظپظ‡ظٹطھظ…ط³ط­ ط¨ط³ ظ…ظ„ظپط§طھ ط§ظ„طھط·ط¨ظٹظ‚ ظ…ظ†ظ‡)' : ''}',
    );
    if (!ok) return;
    var n = 0;
    for (final f in [...outFiles, ...rawFiles]) {
      try {
        f.deleteSync();
        n++;
      } catch (_) {}
    }
    try {
      Directory('$baseDir/.scans').deleteSync(recursive: true);
    } catch (_) {}
    setState(() {
      processedItems.clear();
      missingTerminals = [];
      warnings.clear();
      standardXlsxPath = null;
      generatedPdfPath = null;
      extraPdfPaths = [];
      status = 'طھظ… ظ…ط³ط­ $n ظ…ظ„ظپ.';
    });
  }

  Future<void> openReview() async {
    final list = processedItems.where((e) => e.needsReview).toList();
    if (list.isEmpty) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('ظ…ظپظٹط´ ط¥ظٹطµط§ظ„ط§طھ ظ…ط­طھط§ط¬ط© ظ…ط±ط§ط¬ط¹ط©.')),
      );
      return;
    }
    final changed = await Navigator.push<bool>(
      context,
      MaterialPageRoute(builder: (_) => ReviewPage(items: list, allItems: processedItems)),
    );
    if (changed == true) {
      await rebuildFiles();
    } else if (mounted) {
      setState(() {});
    }
  }

  // ط§ظ„ظ‚ظ„ظ…: ط´ط§ط´ط© ظƒط§ظ…ظ„ط© ظپظٹظ‡ط§ ط§ظ„ط¥ظٹطµط§ظ„ ط¨ط²ظˆظ… + طھط¹ط¯ظٹظ„ ط§ظ„ظ…ط¨ظ„ط؛ ظˆط§ظ„ظ…ط§ظƒظٹظ†ط©
  Future<void> editReceipt(ReceiptItem item) async {
    final changed = await Navigator.push<bool>(
      context,
      MaterialPageRoute(
        builder: (_) => ReviewPage(items: [item], allItems: processedItems, single: true),
      ),
    );
    if (changed == true) {
      await rebuildFiles();
    } else if (mounted) {
      setState(() {});
    }
  }

  void useWhatsApp() {
    saveFolder(null);
    setState(() => customDir = null);
  }

  Future<void> shareStandardFiles() async {
    final files = <XFile>[];
    if (standardXlsxPath != null && File(standardXlsxPath!).existsSync()) {
      files.add(XFile(standardXlsxPath!));
    }
    if (generatedPdfPath != null && File(generatedPdfPath!).existsSync()) {
      files.add(XFile(generatedPdfPath!));
    }
    for (final p in extraPdfPaths) {
      if (File(p).existsSync()) files.add(XFile(p));
    }
    if (files.isNotEmpty) {
      if (!await confirmSend(context,
          review: processedItems.where((e) => e.needsReview).length, missing: missingTerminals.length)) {
        return;
      }
      await Share.shareXFiles(files, text: 'طھظ‚ط±ظٹط± ط§ظ„ط¥ظٹطµط§ظ„ط§طھ (ط§ظ„ط±ط³ظ…ظٹ)');
    } else {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('ظ„ظ… ظٹطھظ… ط¥ظ†ط´ط§ط، ظ…ظ„ظپط§طھ ط¨ط¹ط¯ ظ„ظ…ط´ط§ط±ظƒطھظ‡ط§.')),
      );
    }
  }

  Future<void> shareAdvancedFile() async {
    final files = <XFile>[];
    for (final p in ['$outputDir/review_report.txt', '$outputDir/ocr_debug.txt']) {
      if (File(p).existsSync()) files.add(XFile(p));
    }
    if (files.isNotEmpty) {
      await Share.shareXFiles(files, text: 'ظ„ظˆط¬ ط§ظ„طھظ‚ظپظٹظ„ (طھظ‚ط±ظٹط± ط§ظ„ظ…ط±ط§ط¬ط¹ط© + OCR)');
    } else {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('ظ…ظ„ظپط§طھ ط§ظ„ظ…ط±ط§ط¬ط¹ط© ط؛ظٹط± ظ…طھط§ط­ط© ط­ط§ظ„ظٹط§ظ‹.')),
      );
    }
  }

  Future<void> editReceiptDialog(ReceiptItem item) async {
    final amtController = TextEditingController(text: item.amount.toStringAsFixed(2));
    int? selectedOrder = item.orderNum;

    await showDialog(
      context: context,
      builder: (ctx) => StatefulBuilder(
        builder: (ctx, setDlgState) => AlertDialog(
          title: const Text('طھط¹ط¯ظٹظ„ ط¨ظٹط§ظ†ط§طھ ط§ظ„ط¥ظٹطµط§ظ„'),
          content: SingleChildScrollView(
            child: Column(
              mainAxisSize: MainAxisSize.min,
              children: [
                Image.file(File(item.path), height: 160, fit: BoxFit.contain),
                const SizedBox(height: 8),
                TextField(
                  controller: amtController,
                  keyboardType: const TextInputType.numberWithOptions(decimal: true),
                  decoration: InputDecoration(labelText: tr('ط§ظ„ظ…ط¨ظ„ط؛ ط§ظ„ط¥ط¬ظ…ط§ظ„ظٹ')),
                ),
                const SizedBox(height: 8),
                DropdownButtonFormField<int>(
                  value: selectedOrder,
                  hint: const Text('ط§ط®طھط± ط§ظ„ظ…ط§ظƒظٹظ†ط©'),
                  items: [
                    for (final i in machineNums())
                      DropdownMenuItem(
                        value: i,
                        child: Text(machineLabel(i)),
                      ),
                  ],
                  onChanged: (v) => setDlgState(() => selectedOrder = v),
                ),
              ],
            ),
          ),
          actions: [
            TextButton(onPressed: () => Navigator.pop(ctx), child: const Text('ط¥ظ„ط؛ط§ط،')),
            FilledButton(
              onPressed: () {
                final newAmt = double.tryParse(amtController.text) ?? item.amount;
                setState(() {
                  item.amount = round2(newAmt);
                  item.batches = [item.amount];
                  item.ok = true;
                  item.alt = null;
                  item.how = 'manual';
                  if (selectedOrder != null) {
                    item.orderNum = selectedOrder;
                    item.tid = numToTid[selectedOrder];
                    item.row = terminals[item.tid]!.excelRow;
                    item.unitName = terminals[item.tid]!.unitName;
                  }
                });
                saveFix(item);
                Navigator.pop(ctx);
                rebuildFiles();
              },
              child: const Text('ط­ظپط¸ ظˆطھط­ط¯ظٹط« ط§ظ„ظ…ظ„ظپط§طھ'),
            ),
          ],
        ),
      ),
    );
  }

  Future<void> rebuildFiles() async {
    setState(() => status = 'ط¬ط§ط±ظٹ طھظˆظ„ظٹط¯ ط§ظ„ظ…ظ„ظپط§طھ ظˆطھط­ط¯ظٹط« ط§ظ„ط´ظٹطھط§طھ...');
    final dateCounts = <String, int>{};
    for (final r in processedItems) {
      if (r.date != null) {
        dateCounts[r.date!] = (dateCounts[r.date!] ?? 0) + 1;
      }
    }
    String reportDate = "ط§ظ„ظٹظˆظ…";
    if (dateCounts.isNotEmpty) {
      reportDate = dateCounts.entries.reduce((a, b) => a.value > b.value ? a : b).key;
    }

    // 1. ظ…ظ„ظپط§طھ ط§ظ„ظ€ PDF ط§ظ„ظ…ط¬ظ…ط¹ط©: ظˆط§ط­ط¯ ظ„ظٹظˆظ… ط§ظ„طھظ‚ظپظٹظ„طŒ ظˆظˆط§ط­ط¯ ظ„ظƒظ„ ظٹظˆظ… طھط§ظ†ظٹ ظ„ظˆ ظپظٹظ‡ ط¥ظٹطµط§ظ„ط§طھ ط¨طھط§ط±ظٹط® ظ…ط®طھظ„ظپ
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
        setState(() => status = 'ط¬ط§ط±ظٹ طھط¬ظ‡ظٹط² طµظپط­ط§طھ ط§ظ„ظ€ PDF ($pageNo / ${processedItems.length})...');
        // ط³ظƒط§ظ†ط±: ظ‚طµ ط§ظ„ط¨ط§طھط´ ط¨ط³ + ظپط±ط¯ ط§ظ„ظ…ظ†ط¸ظˆط± + طھط­ط³ظٹظ†. ظ„ظˆ ظپط´ظ„ ظ†ط³طھط®ط¯ظ… ط§ظ„طµظˆط±ط© ط§ظ„ط£طµظ„ظٹط©.
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

    final pdfOutFile = File('$outputDir/ظ†طھظٹط¬ط© $reportDate.pdf');
    await buildPdf(mainItems, pdfOutFile);
    generatedPdfPath = pdfOutFile.path;
    extraPdfPaths = [];
    for (final e in otherByDate.entries) {
      final f = File('$outputDir/ظ†طھظٹط¬ط© ${e.key}.pdf');
      if (await buildPdf(e.value, f)) extraPdfPaths.add(f.path);
    }

    // ط­ط³ط§ط¨ ط¥ط¬ظ…ط§ظ„ظٹط§طھ ظƒظ„ ظˆط­ط¯ط©
    final unitTotals = <int, double>{};
    for (final r in rowNames.keys) {
      unitTotals[r] = 0.0;
    }
    for (final item in mainItems) {
      if (item.row != null && item.amount > 0) {
        unitTotals[item.row!] = (unitTotals[item.row!] ?? 0.0) + item.amount;
      }
    }

    // ط§ظ„ط¥ظƒط³ظٹظ„: ظ†ظپط³ ط§ظ„ظ‚ط§ظ„ط¨ ط¨ط§ظ„ط¸ط¨ط·طŒ ط¨ظ†ظƒطھط¨ ط¨ط³ ط¹ظ…ظˆط¯ D (ظ…ط¨ظ„ط؛ ظƒظ„ ظˆط­ط¯ط©) ظˆ F (ط§ظ„ط¥ط¬ظ…ط§ظ„ظٹ)
    var xlsxLine = 'â€¢ طھظ… طھظˆظ„ظٹط¯ ط§ظ„ط¥ظƒط³ظٹظ„: ظ†طھظٹط¬ط© $reportDate.xlsx';
    standardXlsxPath = null;
    final tpl = await loadTemplateBytes();
    if (tpl == null) {
      xlsxLine = 'âڑ ï¸ڈ ظ‚ط§ظ„ط¨ ط´ظٹطھ ط§ظ„ظ†طھظٹط¬ط© (template.xlsx) ظ…ط´ ظ…ظˆط¬ظˆط¯ ظپظٹ ظپظˆظ„ط¯ط± ط§ظ„طھط·ط¨ظٹظ‚ (ظˆظ„ط§ ظ…ط¶ط§ظپ ظ„ظ„طھط·ط¨ظٹظ‚)طŒ ظپظ…طھط¹ظ…ظ„ط´ ط¥ظƒط³ظٹظ„.';
    } else {
      try {
        final unitRows = rowNames.keys.toList()..sort();
        final outBytes = fillTemplate(tpl, [
          for (final r in unitRows)
            _UnitOut(
              rowNames[r] ?? 'ظˆط­ط¯ط© $r',
              terminals.values.where((t) => t.excelRow == r).length,
              unitTotals[r] ?? 0.0,
            ),
        ]);
        final stdFile = File('$outputDir/ظ†طھظٹط¬ط© $reportDate.xlsx');
        stdFile.createSync(recursive: true);
        stdFile.writeAsBytesSync(outBytes);
        standardXlsxPath = stdFile.path;
      } catch (e) {
        xlsxLine = 'âڑ ï¸ڈ ظپط´ظ„ ظ…ظ„ط، ط§ظ„ظ‚ط§ظ„ط¨: $e';
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

    setState(() => status = 'ط§ظƒطھظ…ظ„!\n'
        '$xlsxLine\n'
        'â€¢ طھظ… طھظˆظ„ظٹط¯ ظ…ظ„ظپ ط§ظ„ظ€ PDF ط§ظ„ظ…ط¬ظ…ط¹.' +
        (extraPdfPaths.isEmpty
            ? ''
            : '\nâ€¢ ظپظٹظ‡ ط¥ظٹطµط§ظ„ط§طھ ط¨طھظˆط§ط±ظٹط® طھط§ظ†ظٹط© (${otherByDate.keys.join(' طŒ ')}): ط§طھط¹ظ…ظ„ ظ„ظ‡ط§ PDF ظ…ظ†ظپطµظ„ ظ„ظƒظ„ ظٹظˆظ…طŒ ظˆظ…ط´ ط¯ط§ط®ظ„ط© ظپظٹ ط¥ط¬ظ…ط§ظ„ظٹط§طھ ظٹظˆظ… $reportDate.') +
        (sourceNote.isEmpty ? '' : '\n$sourceNote'));
  }

  String subtitleOf(ReceiptItem item) {
    final sfx = item.suffix != null ? ' (ط¨ط§طھط´ ${item.suffix})' : '';
    final bt = item.batches.length > 1
        ? '\n= ${item.batches.map((b) => b.toStringAsFixed(2)).join(' + ')}'
        : '';
    final alt = item.alt != null ? '\nâڑ ï¸ڈ ظ‚ط±ط§ط،ط© طھط§ظ†ظٹط©: ${item.alt!.toStringAsFixed(2)}' : '';
    final extra = StringBuffer();
    if (item.ignoredDoc) {
      extra.write('\nًں“„ ط§طھط¬ط§ظ‡ظ„طھظ‡: ط´ظƒظ„ظ‡ ط¥ظٹطµط§ظ„ ط¯ظپط¹/ظ…ط³طھظ†ط¯ ظ…ط´ ط¨ط§طھط´ - ط±ط§ط¬ط¹ظ‡');
    }
    if (item.dateOdd) {
      extra.write('\nًں“… طھط§ط±ظٹط®ظ‡ ${item.date} ظˆط§ظ„ط£ط؛ظ„ط¨ ${item.dateMain}');
    }
    if (!item.ignoredDoc && !item.userExcluded && !item.dupSuspect && item.weakId && item.row != null) {
      extra.write('\nâ‌“ ط§ظ„ظ…ط§ظƒظٹظ†ط©/ط§ظ„ظˆط­ط¯ط© ط§طھط­ط¯ط¯طھ طھظ„ظ‚ط§ط¦ظٹط§ظ‹ - ط§طھط£ظƒط¯');
    }
    final unu = item.unusual != null ? '\nًں“ٹ ${item.unusual}' : '';
    final dup = item.dupSuspect
        ? '\nâڑ ï¸ڈ ظ…ط´ظƒظˆظƒ ط¥ظ†ظ‡ ظ…ظƒط±ط± (ظ†ظپط³ ظ…ط¨ظ„ط؛ ظ…ط§ظƒظٹظ†ط© ${item.dupOf ?? ''}) - ظ…ط´ ظ…ط­ط³ظˆط¨ ظ„ط­ط¯ ظ…ط§ طھط±ط§ط¬ط¹ظ‡'
        : item.userExcluded
            ? '\nًںڑ« ظ…ط³طھط¨ط¹ط¯ (ظ…ظƒط±ط±)'
            : '';
    return 'TID: ${item.tid ?? 'ط؛ظٹط± ظ…ط­ط¯ط¯'}$sfx | ط§ظ„ظ…ط¨ظ„ط؛: ${item.amount.toStringAsFixed(2)} ط¬.ظ…$bt$alt$dup$unu$extra';
  }

  String warningsText() {
    final sb = StringBuffer('ظ…ظ„ط§ط­ط¸ط§طھ ظ„ظ„ظ…ط±ط§ط¬ط¹ط© (${warnings.length}):');
    for (final w in warnings) {
      sb.write('\nâ€¢ $w');
    }
    return sb.toString();
  }

  void writeReport(String date) {
    try {
      final line = '=' * 50;
      final sb = StringBuffer('طھظ‚ط±ظٹط± ط§ظ„ظ…ط±ط§ط¬ط¹ط© - $date\n$line\n');
      var total = 0.0;
      for (final r in processedItems) {
        final flags = <String>[];
        if (r.dupSuspect) flags.add('ظ…ط´ظƒظˆظƒ ط¥ظ†ظ‡ ظ…ظƒط±ط± - ظ…ط´ ظ…ط­ط³ظˆط¨ ظ„ط­ط¯ ظ…ط§ طھط±ط§ط¬ط¹ظ‡');
        if (r.userExcluded) flags.add('ظ…ط³طھط¨ط¹ط¯');
        if (r.ignoredDoc) flags.add('طµظˆط±ط© ظ…طھط¬ط§ظ‡ظ„ط© - ظ…ط­طھط§ط¬ط© ظ‚ط±ط§ط±');
        if (r.dateOdd) flags.add('طھط§ط±ظٹط® ظ…ط®طھظ„ظپ ${r.date}');
        if (r.unusual != null) flags.add('ط¥ط¬ظ…ط§ظ„ظٹ ط¨ط¹ظٹط¯ ط¹ظ† ط§ظ„ظ…ط¹طھط§ط¯');
        if (r.amount <= 0) flags.add('ط§ظ„ظ…ط¨ظ„ط؛ ط؛ظٹط± ظ…ظ‚ط±ظˆط،');
        if (r.row == null) flags.add('ط§ظ„ظˆط­ط¯ط© ط؛ظٹط± ظ…ط¹ط±ظˆظپط©');
        if (!r.ok && r.amount > 0) flags.add('ط§ظ„ظ…ط¨ظ„ط؛ ظٹط­طھط§ط¬ ظ…ط±ط§ط¬ط¹ط©');
        if (r.how == 'auto-number' || r.how == 'conflict' || r.how == 'extra-batch') {
          flags.add('ط±ظ‚ظ… ط§ظ„ظ…ط§ظƒظٹظ†ط©: ${r.how}');
        }
        if (r.batches.length > 1) flags.add('${r.batches.length} ط¨ط§طھط´ط§طھ');
        if (r.alt != null) flags.add('ظ‚ط±ط§ط،ط© طھط§ظ†ظٹط©: ${r.alt!.toStringAsFixed(2)}');
        if (r.amount > 0 && r.row != null && r.counts) total += r.amount;
        final mark = flags.isEmpty ? 'âœ“' : 'âڑ ï¸ڈ';
        final extra = flags.isEmpty ? '' : '   [${flags.join(' | ')}]';
        sb.writeln('$mark ${baseName(r.path)}  ${r.unitName ?? '??'}  '
            '${r.amount.toStringAsFixed(2)}$extra');
      }
      sb.writeln(line);
      sb.writeln('ط§ظ„ط¥ط¬ظ…ط§ظ„ظٹ ط§ظ„ط¹ط§ظ…: ${total.toStringAsFixed(2)} ط¬.ظ…');
      if (missingTerminals.isNotEmpty) {
        sb.writeln('ظ…ط§ظƒظٹظ†ط§طھ ظ†ط§ظ‚طµط©: ${missingTerminals.join(' طŒ ')}');
      }
      for (final w in warnings) {
        sb.writeln('- $w');
      }
      File('$outputDir/review_report.txt')
        ..createSync(recursive: true)
        ..writeAsStringSync(sb.toString());
    } catch (_) {}
  }

  Future<void> runFullPipeline() async {
    try {
      await Permission.notification.request();
    } catch (_) {}
    try {
      await _chan.invokeMethod('startKeepAlive',
          {'title': 'A2R POS', 'text': 'ط¬ط§ط±ظٹ ظپط­طµ ط§ظ„ط¥ظٹطµط§ظ„ط§طھ...'});
    } catch (_) {}
    try {
      await _chan.invokeMethod('keepScreenOn', {'on': true});
    } catch (_) {}
    try {
      await _runPipelineInner();
    } finally {
      try {
        await _chan.invokeMethod('stopKeepAlive');
      } catch (_) {}
      try {
        await _chan.invokeMethod('keepScreenOn', {'on': false});
      } catch (_) {}
      if (_life != AppLifecycleState.resumed) {
        final rv = processedItems.where((e) => e.needsReview).length;
        final txt = 'طھظ… ظپط­طµ ${processedItems.length} ط¥ظٹطµط§ظ„'
            '${missingTerminals.isNotEmpty ? ' - ظ†ط§ظ‚طµ ${missingTerminals.length} ظ…ط§ظƒظٹظ†ط©' : ''}'
            '${rv > 0 ? ' - $rv ظ„ظ„ظ…ط±ط§ط¬ط¹ط©' : ''}';
        try {
          await _chan.invokeMethod('notify', {'title': tr('A2R POS - ط§ظ†طھظ‡ظ‰ ط§ظ„ظپط­طµ'), 'text': tr(txt)});
        } catch (_) {}
      }
    }
  }

  Future<void> _runPipelineInner({List<String>? manualPaths}) async {
    final perm = await Permission.manageExternalStorage.request();
    if (!perm.isGranted) {
      setState(() => status = 'ظٹط¬ط¨ ظ…ظ†ط­ ط¥ط°ظ† ط§ظ„ظˆطµظˆظ„ ط¥ظ„ظ‰ ط¬ظ…ظٹط¹ ط§ظ„ظ…ظ„ظپط§طھ ظ„ظ„ظ…طھط§ط¨ط¹ط©.');
      return;
    }
    await loadSettingsFromExcel();

    setState(() {
      busy = true;
      progress = 0;
      processedItems.clear();
      missingTerminals.clear();
      warnings.clear();
      excludedPay.clear();
      status = 'ط¬ط§ط±ظٹ ط§ظ„ط¨ط­ط« ط¹ظ† ظ…ط¬ظ„ط¯ط§طھ ط§ظ„طµظˆط±...';
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
          if (!(lower.endsWith('.jpg') ||
              lower.endsWith('.jpeg') ||
              lower.endsWith('.png') ||
              lower.endsWith('.webp'))) {
            continue;
          }
          if (baseName(e.path).startsWith('.') || baseName(e.path).startsWith('dup_')) continue;
          if (fileDate(e).isBefore(cutoff)) continue;
          todo.add(e);
        }
      } catch (_) {}
    }
    todo.sort((a, b) => baseName(a.path).compareTo(baseName(b.path)));

    sourceNote = '';
    if (manualPaths != null) {
      todo.addAll(manualPaths.map((p) => File(p)));
      sourceNote = 'â€¢ ط§ظ„ظ…طµط¯ط±: طµظˆط± ظ…ط­ط¯ط¯ط© ظٹط¯ظˆظٹط§ظ‹ (${todo.length} طµظˆط±ط©)';
    }

    // ظپظˆظ„ط¯ط± ط§ط®طھط±طھظ‡ ط¨ط¥ظٹط¯ظƒ (ظ…ط´ ظˆط§طھط³ط§ط¨): ط£ظٹ ظ…ظ„ظپ PDF ط¬ظˆط§ظ‡ ط¨ظٹطھظپط­طµطŒ ط¨ط£ظٹ ط§ط³ظ… ظˆط¨ط£ظٹ طھط§ط±ظٹط®
    if (customDir != null && customDir != outputDir) {
      final cpdfs = <File>[];
      try {
        for (final e in Directory(customDir!).listSync().whereType<File>()) {
          final n = baseName(e.path);
          if (!n.startsWith('.') && n.toLowerCase().endsWith('.pdf')) cpdfs.add(e);
        }
      } catch (_) {}
      cpdfs.sort((a, b) => baseName(a.path).compareTo(baseName(b.path)));
      var pdfPages = 0;
      final used = <String>[];
      for (final pf in cpdfs) {
        setState(() => status = 'ط¨ظ‚ط±ط£ طµظپط­ط§طھ ظ…ظ„ظپ PDF: ${baseName(pf.path)}...');
        final pages = await extractPdfPages(pf);
        if (pages.isEmpty) continue;
        todo.addAll(pages);
        pdfPages += pages.length;
        used.add(baseName(pf.path));
      }
      if (used.isNotEmpty) {
        sourceNote = 'â€¢ ط§ظ„ظ…طµط¯ط±: ظ…ظ„ظپ PDF ${used.take(3).join(' طŒ ')}${used.length > 3 ? ' ...' : ''} ($pdfPages طµظپط­ط©)';
      }
    }

    // ظ…ظپظٹط´ طµظˆط±طں ط¯ظˆظ‘ط± ط¹ظ„ظ‰ PDF طھظ‚ظپظٹظ„ ط¨طھط§ط±ظٹط® ط§ظ„ظ†ظ‡ط§ط±ط¯ظ‡
    if (todo.isEmpty) {
      setState(() => status = 'ظ…ظپظٹط´ طµظˆط±. ط¨ط¯ظˆظ‘ط± ط¹ظ„ظ‰ ظ…ظ„ظپ PDF طھظ‚ظپظٹظ„ ط¨طھط§ط±ظٹط® ${todayToken()}...');
      final searchDirs = <String>[];
      if (customDir != null) searchDirs.add(customDir!);
      for (final d in findWhatsAppImageDirs()) {
        final p = d.path;
        if (p.endsWith('WhatsApp Images')) {
          searchDirs.add(p.substring(0, p.length - 'WhatsApp Images'.length) + 'WhatsApp Documents');
        }
      }
      searchDirs.addAll([baseDir, '$root/Download', '$root/Documents']);
      final pdfs = findTodayAlexPdfs(searchDirs);
      if (pdfs.isNotEmpty) {
        final pages = await extractPdfPages(pdfs.first);
        if (pages.isNotEmpty) {
          todo.addAll(pages);
          sourceNote = 'â€¢ ط§ظ„ظ…طµط¯ط±: ظ…ظ„ظپ PDF ${baseName(pdfs.first.path)} (${pages.length} طµظپط­ط©)';
        }
      }
    }

    if (todo.isEmpty) {
      setState(() {
        busy = false;
        status = customDir != null
            ? 'ظ…ظپظٹط´ طµظˆط± (ظپظٹ ط§ظ„ظپطھط±ط© ط§ظ„ظ…ط­ط¯ط¯ط©) ظˆظ„ط§ ظ…ظ„ظپط§طھ PDF ظپظٹ ط§ظ„ظ…ط¬ظ„ط¯ ط§ظ„ظ…ط®طھط§ط±.'
            : (dirs.isEmpty
                    ? 'طھط¹ط°ط± ط§ظ„ط¹ط«ظˆط± ط¹ظ„ظ‰ ظ…ط¬ظ„ط¯ طµظˆط± ط§ظ„ظˆط§طھط³ط§ط¨ طھظ„ظ‚ط§ط¦ظٹط§ظ‹.'
                    : 'ظ„ط§ طھظˆط¬ط¯ طµظˆط± ظپظٹ ظ…ط¬ظ„ط¯ ط§ظ„ظˆط§طھط³ط§ط¨ ط¨ط§ظ„ظپطھط±ط© ط§ظ„ظ…ط­ط¯ط¯ط©.') +
                '\nظˆظ„ط§ ظٹظˆط¬ط¯ ظ…ظ„ظپ PDF طھظ‚ظپظٹظ„ ط¨طھط§ط±ظٹط® ${todayToken()}.';
      });
      return;
    }

    final recognizer = TextRecognizer(script: TextRecognitionScript.latin);
    Directory(rawImagesDir).createSync(recursive: true);
    Directory(outputDir).createSync(recursive: true);
    try {
      File('$outputDir/ocr_debug.txt').writeAsStringSync('');
    } catch (_) {}
    final fixes = loadFixes();
    final rawReceipts = <ReceiptItem>[];
    final ignoredItems = <ReceiptItem>[];

    // â•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گ ط§ظ„ظ…ط±ط­ظ„ط© 1: ظ‚ط±ط§ط،ط© ط§ظ„طµظˆط± â•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گ
    for (var i = 0; i < todo.length; i++) {
      final f = todo[i];
      try {
        Directory(tmpDir).createSync(recursive: true);
        var workPath = f.path;
        var res = await recognizer.processImage(InputImage.fromFilePath(workPath));
        final fkey = '${baseName(f.path)}:${f.lengthSync()}';
        final fxm = fixes[fkey];
        final ign = (fxm is Map && fxm['ign'] is String) ? fxm['ign'] as String : null;
        var isReceipt = receiptScore(res.text) >= 2 || findTerminalIn([res.text]) != null;
        if (ign == 'keep') isReceipt = true; // ظ‚ط±ط±طھ ظ‚ط¨ظ„ ظƒط¯ظ‡ ط¥ظ†ظ‡ط§ ط¨ط§طھط´
        // ط¥ظٹطµط§ظ„ ط¯ظپط¹ ط£ظˆ ظ…ط³طھظ†ط¯ ظ…ط´ ط¨ط§طھط´/ظ…ظ„ط®طµ/ط±ظٹ ط¨ط±ظ†طھ: ط¨ظٹط¯ط®ظ„ ط§ظ„ظ…ط±ط§ط¬ط¹ط© (ظ…ط´ ط¨ظٹطھطھط¬ط§ظ‡ظ„ ظ†ظ‡ط§ط¦ظٹ)
        var skipPay = false;
        if (isReceipt && ign != 'keep' && classifyReceipt(res.text) == 'payment') {
          isReceipt = false;
          skipPay = true;
        }

        // طµظˆط±ط© ظ…ظ‚ظ„ظˆط¨ط©طں ط¬ط±ظ‘ط¨ EXIF ط«ظ… 90 / 270 / 180
        if (!isReceipt && !skipPay && autoRotate) {
          for (final deg in const [0, 90, 270, 180]) {
            final tp = '$tmpDir/rot_$deg.jpg';
            try {
              final wrote = await fixImageAsync(f.path, tp, deg);
              if (!wrote) continue;
              final r2 = await recognizer.processImage(InputImage.fromFilePath(tp));
              if ((receiptScore(r2.text) >= 2 || findTerminalIn([r2.text]) != null) &&
                  classifyReceipt(r2.text) != 'payment') {
                res = r2;
                workPath = tp;
                isReceipt = true;
                break;
              }
            } catch (_) {}
          }
        }

        if (isReceipt) {
          final key = '${baseName(f.path)}:${f.lengthSync()}';
          final a = analyzeLines(toLines(res));

          // ظ‚ط±ط§ط،ط© طھط§ظ†ظٹط© ط¨طھظ†ط³ظٹظ‚ ظ…ط®طھظ„ظپ (طµظˆط±ط© ظ…ط­ط³ظ‘ظ†ط©) ط¹ظ„ط´ط§ظ† ظ†ظƒط´ظپ ط§ظ„ط£ط®ط·ط§ط،
          double? alt;
          // ط§ظ„ظ‚ط±ط§ط،ط© ط§ظ„طھط§ظ†ظٹط© ط¨ط³ ظ„ظ„ط¥ظٹطµط§ظ„ط§طھ ط§ظ„ظ„ظٹ ط§ظ„ظ‚ط±ط§ط،ط© ط§ظ„ط£ظˆظ„ظ‰ ظپظٹظ‡ط§ ظ…ط´ ظ…ط¤ظƒط¯ط© (ط¹ظ„ط´ط§ظ† ط§ظ„ط³ط±ط¹ط©)
          final needSecond = !a.ok || a.tid == null || a.row == null || a.val <= 0;
          if (needSecond) {
            try {
            final ep = '$tmpDir/enh.jpg';
            await enhanceImageAsync(workPath, ep);
            final res2 = await recognizer.processImage(InputImage.fromFilePath(ep));
            final a2 = analyzeLines(toLines(res2));
            try {
              File('$outputDir/ocr_debug.txt').writeAsStringSync(
                  '##### ${baseName(f.path)} (2nd read) total=${a2.val} ok=${a2.ok}\n',
                  mode: FileMode.append);
            } catch (_) {}
            alt = mergeSecondRead(a, a2);
            } catch (_) {}
          }

          if (a.tid == null) {
            final fnNum = int.tryParse(baseName(f.path).split('.').first);
            final ft = fnNum != null ? numToTid[fnNum] : null;
            if (ft != null && (a.row == null || terminals[ft]!.excelRow == a.row)) {
              a.tid = ft;
              a.row = terminals[ft]!.excelRow;
              a.unitName = terminals[ft]!.unitName;
              a.how = 'file-number';
            }
          }
          try {
            File('$outputDir/ocr_debug.txt').writeAsStringSync(
                '##### ${baseName(f.path)} => tid=${a.tid} how=${a.how} total=${a.val} ok=${a.ok} alt=$alt rotated=${workPath != f.path}\n'
                '${a.debug}\n--- raw ML Kit ---\n${res.text}\n\n',
                mode: FileMode.append);
          } catch (_) {}

          // ظ†ط³ط®ط© ط§ظ„طµظˆط±ط© ط§ظ„ظ„ظٹ ظ‡طھط¯ط®ظ„ ط§ظ„ظ€ PDF (ظ…طھط¸ط¨ط·ط© ط§ظ„ط§طھط¬ط§ظ‡)
          final inRaw = f.path.startsWith('$rawImagesDir/');
          final dst = inRaw ? f.path : '$rawImagesDir/wa_${baseName(f.path)}';
          File localFile;
          if (workPath != f.path) {
            final noExt = dst.contains('.') ? dst.substring(0, dst.lastIndexOf('.')) : dst;
            final jd = '$noExt.jpg';
            File(workPath).copySync(jd);
            if (jd != dst && inRaw) {
              try {
                File(dst).deleteSync();
              } catch (_) {}
            }
            localFile = File(jd);
          } else if (inRaw || File(dst).existsSync()) {
            localFile = File(dst);
          } else {
            await fixImageAsync(f.path, dst, 0); // ظ†ط³ط® + طھطµط­ظٹط­ EXIF ظ„ظˆ ظ„ط§ط²ظ…
            localFile = File(dst);
          }

          final item = ReceiptItem(
            path: localFile.path,
            tid: a.tid,
            orderNum: a.tid != null ? terminals[a.tid]!.orderNum : null,
            row: a.row,
            unitName: a.unitName,
            amount: a.val,
            date: a.date,
            key: key,
            ok: a.ok,
            how: a.how,
            batches: a.batches,
            alt: alt,
          );
          item.txn = a.txn;

          // طھطµط­ظٹط­ ظٹط¯ظˆظٹ ظ…ط­ظپظˆط¸ ظ…ظ† ظ…ط±ط© ط³ط§ط¨ظ‚ط©
          final fx = fixes[key];
          if (fx is Map) {
            final v = fx['val'];
            if (v is num && v > 0) {
              item.amount = round2(v.toDouble());
              item.batches = [item.amount];
            }
            final ft = fx['tid'];
            if (ft is String && terminals.containsKey(ft)) {
              final t = tidAlias[ft] ?? ft;
              final info = terminals[t]!;
              item.tid = t;
              item.orderNum = info.orderNum;
              item.row = info.excelRow;
              item.unitName = info.unitName;
            }
            item.ok = true;
            item.alt = null;
            item.how = 'manual';
            final dd = fx['dup'];
            if (dd is String) item.dupDecision = dd;
            final ig = fx['ign'];
            if (ig is String) item.ign = ig;
            final fd = fx['date'];
            if (fd is String) {
              item.date = fd;
              item.dateFix = fd;
            }
          }
          rawReceipts.add(item);
        } else if (skipPay) {
          excludedPay.add(baseName(f.path));
          if (ign != 'drop') {
            ignoredItems.add(ReceiptItem(
              path: f.path,
              key: fkey,
              how: 'none',
              ignoredDoc: true,
            ));
          }
        }
      } catch (e) {
        warnings.add('${baseName(f.path)}: ظپط´ظ„طھ ط§ظ„ظ‚ط±ط§ط،ط© (${e.toString().split('\n').first})');
      }

      if (!mounted) return;
      setState(() {
        progress = ((i + 1) / todo.length) * 0.5;
        status = 'ظپط­طµ OCR: ${i + 1} ظ…ظ† ${todo.length}';
      });
    }

    await recognizer.close();

    if (rawReceipts.isEmpty) {
      processedItems.addAll(ignoredItems);
      setState(() {
        busy = false;
        status = 'طھظ… ظپط­طµ ${todo.length} طµظˆط±ط©طŒ ظˆظ„ظ… ظٹطھظ… ط§ظ„ط¹ط«ظˆط± ط¹ظ„ظ‰ ط¥ظٹطµط§ظ„ط§طھ ظ…ط·ط§ط¨ظ‚ط©.';
      });
      return;
    }

    // â•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گ ط§ظ„ظ…ط±ط­ظ„ط© 2: ط§ظ„طھظƒط±ط§ط± ظˆطھط±ظ‚ظٹظ… ط§ظ„ظ…ط§ظƒظٹظ†ط§طھ â•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گ
    setState(() => status = 'ظ…ط±ط§ط¬ط¹ط© ط§ظ„طھظƒط±ط§ط± ظˆطھط±ظ‚ظٹظ… ط§ظ„ظ…ط§ظƒظٹظ†ط§طھ...');
    final prio = <String, int>{'manual': 0, 'terminal': 0, 'fuzzy': 1, 'file-number': 2};
    final withTid = rawReceipts.where((r) => r.tid != null).toList();
    withTid.sort((a, b) {
      final c = (prio[a.how] ?? 3).compareTo(prio[b.how] ?? 3);
      return c != 0 ? c : rawReceipts.indexOf(a).compareTo(rawReceipts.indexOf(b));
    });

    final taken = <int, ReceiptItem>{};
    for (final r in withTid) {
      final n = r.orderNum!;
      final prev = taken[n];
      if (prev == null) {
        taken[n] = r;
      } else if (r.amount > 0 &&
          (prev.amount - r.amount).abs() < 0.005 &&
          sameDay(prev, r)) {
        if (r.dupDecision == 'drop') {
          r.userExcluded = true; // ط§ظ†طھ ظ‚ط±ط±طھ ظ‚ط¨ظ„ ظƒط¯ظ‡ ط¥ظ†ظ‡ ظ…ظƒط±ط± (ط¨ظٹظپط¶ظ„ ط¸ط§ظ‡ط± ظˆطھظ‚ط¯ط± طھط±ط¬ظ‘ط¹ظ‡)
          r.dupOf = n;
          warnings.add('${baseName(r.path)}: ط¥ظٹطµط§ظ„ ظ…ظƒط±ط± ظ„ظ„ظ…ط§ظƒظٹظ†ط© $n (ط§ط³طھط¨ط¹ط¯طھظ‡ ظ‚ط¨ظ„ ظƒط¯ظ‡ - طھظ‚ط¯ط± طھط±ط¬ظ‘ط¹ظ‡)');
        } else if (r.dupDecision == 'keep') {
          r.origTid = r.tid; // ط§ظ†طھ ظ‚ط±ط±طھ ظ‚ط¨ظ„ ظƒط¯ظ‡ ط¥ظ†ظ‡ ط¨ط§طھط´ طھط§ظ†ظٹ ط­ظ‚ظٹظ‚ظٹ
          r.tid = null;
          r.conflict = true;
        } else {
          // ظ†ظپط³ ط§ظ„ظ…ط§ظƒظٹظ†ط© ظˆظ†ظپط³ ط§ظ„ظ…ط¨ظ„ط؛ = ظ…ط´ظƒظˆظƒ ط¥ظ†ظ‡ ظ…ظƒط±ط±: ط¨ظٹط¯ط®ظ„ ط§ظ„ظ…ط±ط§ط¬ط¹ط© (ظ…ط´ ط¨ظٹطھط³طھط¨ط¹ط¯)
          r.dupSuspect = true;
          r.dupOf = n;
          warnings.add('${baseName(r.path)}: ظ†ظپط³ ط§ظ„ظ…ط§ظƒظٹظ†ط© $n ظˆظ†ظپط³ ط§ظ„ظ…ط¨ظ„ط؛ '
              '${r.amount.toStringAsFixed(2)} (ظ…ط´ظƒظˆظƒ ط¥ظ†ظ‡ ظ…ظƒط±ط±) - ط±ط§ط¬ط¹ظ‡ ظ…ظ† ط²ط± ط§ظ„ظ…ط±ط§ط¬ط¹ط©');
        }
      } else {
        r.origTid = r.tid;
        r.tid = null;
        r.conflict = true;
      }
    }

    final usedSuffix = <String>{};
    void giveSuffix(ReceiptItem r) {
      var k = 2;
      while (usedSuffix.contains('${r.origTid}:$k')) {
        k++;
      }
      usedSuffix.add('${r.origTid}:$k');
      r.tid = r.origTid;
      r.suffix = k;
      r.how = 'extra-batch';
    }

    for (final r in rawReceipts) {
      if (r.skip || r.tid != null || r.row == null) continue;
      if (r.conflict &&
          (r.how == 'terminal' || r.how == 'fuzzy' || r.dupDecision == 'keep')) {
        giveSuffix(r); // ظ†ظپط³ ط§ظ„ظ…ط§ظƒظٹظ†ط© ط¨ظ…ط¨ظ„ط؛ ظ…ط®طھظ„ظپ = ط¨ط§طھط´ ط¥ط¶ط§ظپظٹ
        continue;
      }
      final free = <int>[];
      for (final n in machineNums()) {
        if (terminals[numToTid[n]!]!.excelRow == r.row && !taken.containsKey(n)) {
          free.add(n);
        }
      }
      if (free.isNotEmpty) {
        final n = free.first;
        final t = numToTid[n]!;
        r.tid = t;
        r.orderNum = n;
        r.unitName = terminals[t]!.unitName;
        taken[n] = r;
        r.how = r.conflict ? 'conflict' : 'auto-number';
      } else if (r.origTid != null) {
        giveSuffix(r);
      }
    }

    final kept = rawReceipts.where((r) => !r.skip).toList();

    // ط§ظ„طھط§ط±ظٹط® ط§ظ„ط£ظƒط«ط± طھظƒط±ط§ط±ط§ظ‹ + طھط­ط°ظٹط±ط§طھ
    final dateVotes = <String, int>{};
    for (final r in kept) {
      if (r.date != null) dateVotes[r.date!] = (dateVotes[r.date!] ?? 0) + 1;
    }
    final mainDate = mostCommon(dateVotes.keys.expand((k) => List.filled(dateVotes[k]!, k)).toList())?.key;

    for (final r in kept) {
      final nm = baseName(r.path);
      if (mainDate != null && r.date != null && r.date != mainDate && r.dateFix == null) {
        r.dateOdd = true; // ط¨ظٹط¯ط®ظ„ ط§ظ„ظ…ط±ط§ط¬ط¹ط©
        r.dateMain = mainDate;
      }
      if (r.row != null && r.amount > 0) {
        if (!r.ok) warnings.add('$nm: ${r.unitName} = ${r.amount.toStringAsFixed(2)} (ط±ط§ط¬ط¹ ط§ظ„ظ…ط¨ظ„ط؛ ظٹط¯ظˆظٹط§ظ‹)');
        if (r.alt != null) {
          warnings.add('$nm: ط§ظ„ظ‚ط±ط§ط،ط© ط§ظ„طھط§ظ†ظٹط© ظ„ظ„ظ…ط¨ظ„ط؛ ${r.alt!.toStringAsFixed(2)} ظ…ط®طھظ„ظپط© ط¹ظ† ${r.amount.toStringAsFixed(2)} - ط±ط§ط¬ط¹ ط§ظ„طµظˆط±ط©');
        }
        if (r.how == 'extra-batch') {
          warnings.add('$nm: ط§ظ„ظ…ط§ظƒظٹظ†ط© ${r.orderNum} ط¸ظ‡ط±طھ ظپظٹ طµظˆط±ط© طھط§ظ†ظٹط© ط¨ظ…ط¨ظ„ط؛ ظ…ط®طھظ„ظپ - ط§طھط­ط³ط¨طھ ظƒط¨ط§طھط´ ط¥ط¶ط§ظپظٹ ${r.suffix}');
        } else if (r.how == 'auto-number' || r.how == 'conflict') {
          warnings.add('$nm: ط±ظ‚ظ… ط§ظ„ظ…ط§ظƒظٹظ†ط© ط§طھط­ط¯ط¯ طھظ„ظ‚ط§ط¦ظٹط§ظ‹ (${r.orderNum}) - ط§طھط£ظƒط¯ ظ…ظ†ظ‡');
        } else if (!(r.how == 'terminal' || r.how == 'fuzzy' || r.how == 'manual')) {
          warnings.add('$nm: ط§ظ„ظˆط­ط¯ط© ط§طھط­ط¯ط¯طھ ط¨ط§ظ„ط§ط³ظ…/ط§ظ„ظ…ظٹط±ط´ط§ظ†طھ ظپظ‚ط·');
        }
      } else {
        warnings.add('$nm: ظ…ط´ ظ‚ط§ط¯ط± ط£ظ‚ط±ط§ ط§ظ„ظ…ط¨ظ„ط؛ ط£ظˆ ط§ظ„ظˆط­ط¯ط© - ط§ط¶ط؛ط· طھط¹ط¯ظٹظ„');
      }
    }

    // â•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گ ط§ظ„ظ…ط±ط­ظ„ط© 3: ط¥ط¹ط§ط¯ط© ط§ظ„طھط³ظ…ظٹط© (ط¹ظ„ظ‰ ظ…ط±ط­ظ„طھظٹظ† ط¹ط´ط§ظ† ط§ظ„طھط¹ط§ط±ط¶) â•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گ
    setState(() => status = 'ط¥ط¹ط§ط¯ط© طھط³ظ…ظٹط© ط§ظ„طµظˆط± ظˆطھط±طھظٹط¨ظ‡ط§...');
    final moves = <ReceiptItem>[];
    final newNames = <ReceiptItem, String>{};
    var dupN = 0;
    for (final r in kept) {
      if (r.tid == null || r.orderNum == null) continue;
      final ext = r.path.split('.').last;
      final sfx = r.suffix != null
          ? '_${r.suffix}'
          : ((r.dupSuspect || r.userExcluded) ? '_dup${++dupN}' : '');
      final target = '$rawImagesDir/${r.orderNum}$sfx.$ext';
      if (r.path != target) {
        moves.add(r);
        newNames[r] = target;
      }
    }
    for (var k = 0; k < moves.length; k++) {
      final r = moves[k];
      final tmp = '$rawImagesDir/__tmp$k.${r.path.split('.').last}';
      try {
        r.path = File(r.path).renameSync(tmp).path;
      } catch (_) {}
    }
    for (final r in moves) {
      try {
        final dst = newNames[r]!;
        if (File(dst).existsSync()) File(dst).deleteSync();
        r.path = File(r.path).renameSync(dst).path;
      } catch (_) {}
    }

    processedItems.addAll(kept);
    processedItems.addAll(ignoredItems);
    await rebuildFiles();
    try {
      Directory(tmpDir).deleteSync(recursive: true);
    } catch (_) {}

    if (!mounted) return;
    setState(() {
      busy = false;
      progress = 1.0;
    });
  }

  Widget _card(String title, List<Widget> children) => Card(
        margin: const EdgeInsets.only(bottom: 10),
        child: Padding(
          padding: const EdgeInsets.all(12),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.stretch,
            children: [
              Text(title, style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 15)),
              const SizedBox(height: 8),
              ...children,
            ],
          ),
        ),
      );

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
          PopupMenuButton<String>(
            tooltip: tr('ط§ظ„ظ…ط²ظٹط¯'),
            onSelected: (v) {
              if (v == 'appearance') {
                showAppearanceSheet(context);
              } else if (v == 'log') {
                shareDebugLog(context);
              } else if (v == 'folder') {
                openFolder(outputDir);
              } else if (v == 'src') {
                useWhatsApp();
              } else if (v == 'out') {
                useDefaultOut();
              } else if (v == 'settings') {
                ensureSettingsFile().then((ok) {
                  if (!mounted) return;
                  if (ok) {
                    openFile(settingsPath, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                  } else {
                    ScaffoldMessenger.of(context).showSnackBar(
                      const SnackBar(content: Text('ظ…ظ‚ط¯ط±طھط´ ط£ط¹ظ…ظ„ ظ…ظ„ظپ ط§ظ„ط¥ط¹ط¯ط§ط¯ط§طھ. ط§طھط£ظƒط¯ ظ…ظ† ط¥ط°ظ† ط§ظ„ظˆطµظˆظ„ ظ„ظƒظ„ ط§ظ„ظ…ظ„ظپط§طھ.')),
                    );
                  }
                });
              } else if (v == 'wipe') {
                wipeEverything(context);
              } else if (v == 'clear') {
                clearAll();
              }
            },
            itemBuilder: (_) => [
              if (MediaQuery.sizeOf(context).width < 400)
                PopupMenuItem(
                    value: 'appearance',
                    child: w.Text(AppSettings.of(context).isAr ? 'ط§ظ„ظ…ط¸ظ‡ط± ظˆط§ظ„ظ„ط؛ط©' : 'Theme & language')),
              const PopupMenuItem(value: 'folder', child: Text('ظپطھط­ ظپظˆظ„ط¯ط± ط§ظ„ط¥ط®ط±ط§ط¬')),
              if (customDir != null)
                PopupMenuItem(value: 'src', enabled: !busy, child: const Text('ط§ظ„ظ…طµط¯ط±: ط±ط¬ظˆط¹ ظ„ظˆط§طھط³ط§ط¨')),
              if (customOut != null)
                PopupMenuItem(value: 'out', enabled: !busy, child: const Text('ط§ظ„ط¥ط®ط±ط§ط¬: ط§ظ„ط§ظپطھط±ط§ط¶ظٹ')),
              PopupMenuItem(
                  value: 'log',
                  child: w.Text(AppSettings.of(context).isAr ? 'ظ…ط´ط§ط±ظƒط© ط§ظ„ظ„ظˆط¬' : 'Share log')),
              const PopupMenuItem(enabled: false, child: Text('ط§ظ„ط¥طµط¯ط§ط± $kAppVersion')),
              PopupMenuItem(value: 'settings', child: const Text('ط¥ط¹ط¯ط§ط¯ط§طھ ط§ظ„ظ…ط§ظƒظٹظ†ط§طھ (ط¥ظƒط³ظٹظ„)')),
              PopupMenuItem(value: 'clear', enabled: !busy, child: const Text('ظ…ط³ط­ ط§ظ„ظƒظ„')),
              PopupMenuItem(
                  value: 'wipe',
                  enabled: !busy,
                  child: w.Text(AppSettings.of(context).isAr ? 'ًں§¹ ظ…ط³ط­ ظƒظ„ ط­ط§ط¬ط© (ظƒط§ط´ + ظ…ط¤ظ‚طھ)' : 'ًں§¹ Wipe everything (cache + temp)')),
            ],
          ),
        ],
      ),
      bottomNavigationBar: _bottomBar(
        FilledButton.icon(
          onPressed: busy ? null : runFullPipeline,
          icon: const Icon(Icons.play_arrow),
          label: const Text('ط¨ط¯ط، ط§ظ„ظ…ط¹ط§ظ„ط¬ط© ط§ظ„ط´ط§ظ…ظ„ط©'),
        ),
        busy: busy,
        progress: progress,
      ),
      body: ListView(
        padding: const EdgeInsets.fromLTRB(8, 6, 8, 8),
        children: [
          _settingsCard([
            _periodRow(days, busy, (d) {
              setState(() => days = d);
              sharedDays.value = d;
            }),
            const Divider(height: 10),
            _pathRow(Icons.folder_open, 'ط§ظ„ظ…طµط¯ط±', customDir ?? 'ظˆط§طھط³ط§ط¨ (طھظ„ظ‚ط§ط¦ظٹ)',
                busy ? null : pickFolder,
                onReset: customDir != null && !busy ? useWhatsApp : null),
            _pathRow(Icons.add_photo_alternate, 'ط§ط®طھظٹط§ط± طµظˆط± ظٹط¯ظˆظٹط§ظ‹', 'ظپطھط­ ظ…ط¹ط±ط¶ ط§ظ„طµظˆط±', busy ? null : pickManualImages),
            _pathRow(Icons.drive_file_move_outline, 'ط§ظ„ط¥ط®ط±ط§ط¬',
                customOut ?? '$defaultOutputDir (ط§ظپطھط±ط§ط¶ظٹ)', busy ? null : pickOutFolder,
                onReset: customOut != null && !busy ? useDefaultOut : null),
            _switchRow('ظپط­طµ ط§ظ„طµظˆط± ط§ظ„ظ…ظ‚ظ„ظˆط¨ط© (ط£ط¨ط·ط£)', autoRotate,
                busy ? null : (v) => setState(() => autoRotate = v)),
          ]),
          _statusBox(status),
          const SizedBox(height: 6),
          if (processedItems.isNotEmpty) ...[
            _filterRow(reviewOnly, processedItems.length, reviewCount, (v) => setState(() => reviewOnly = v)),
            if (reviewCount > 0)
              Padding(
                padding: const EdgeInsets.only(top: 4),
                child: w.Text(
                    _l(context, 'ًں’، ط§ط³ط­ط¨ ط¥ظٹطµط§ظ„ ط§ظ„ظ…ط±ط§ط¬ط¹ط©: طµط­ âœ“ ظپظٹ ظ†ط§ط­ظٹط© ظˆط؛ظ„ط· âœ— ظپظٹ ط§ظ„ظ†ط§ط­ظٹط© ط§ظ„طھط§ظ†ظٹط©',
                        'ًں’، Swipe a review receipt: OK âœ“ one way, wrong âœ— the other'),
                    style: TextStyle(fontSize: 11, color: Theme.of(context).colorScheme.onSurfaceVariant)),
              ),
            const SizedBox(height: 6),
          ],
          Wrap(
            spacing: 6,
            runSpacing: 6,
            children: [
              _mini(
                  Icons.table_chart,
                  'ط¥ظƒط³ظٹظ„',
                  xlsx == null
                      ? null
                      : () => openFile(xlsx,
                          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
                  filled: true),
              _mini(Icons.picture_as_pdf, 'PDF',
                  pdf == null ? null : () => openFile(pdf, 'application/pdf'),
                  filled: true),
              _mini(Icons.chat, 'ظˆط§طھط³ط§ط¨', busy ? null : sendToWhatsApp),
              _mini(Icons.fact_check, 'ظ…ط±ط§ط¬ط¹ط© ($reviewCount)', busy ? null : openReview),
            ],
          ),
          if (missingTerminals.isNotEmpty) ...[
            const SizedBox(height: 10),
            Container(
              padding: const EdgeInsets.all(10),
              decoration: BoxDecoration(
                color: _tintBg(context, Colors.amber),
                border: Border.all(color: _tintBorder(context, Colors.amber)),
                borderRadius: BorderRadius.circular(8),
              ),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Text(
                    'âڑ ï¸ڈ ظ…ط§ظƒظٹظ†ط§طھ ظ†ط§ظ‚طµط© (${missingTerminals.length} ظ…ظ† ${terminals.length}):',
                    style: TextStyle(
                      fontWeight: FontWeight.bold,
                      color: _tintFg(context, Colors.amber),
                    ),
                  ),
                  const SizedBox(height: 4),
                  Text(
                    missingTerminals
                        .map((n) => '$n (${terminals[numToTid[n]]?.unitName ?? ''})')
                        .join(' طŒ '),
                    style: const TextStyle(fontSize: 13),
                  ),
                ],
              ),
            ),
          ],
          if (warnings.isNotEmpty) ...[
            const SizedBox(height: 10),
            Container(
              constraints: const BoxConstraints(maxHeight: 130),
              padding: const EdgeInsets.all(8),
              decoration: BoxDecoration(
                color: _tintBg(context, Colors.orange, soft: true),
                border: Border.all(color: _tintBorder(context, Colors.orange)),
                borderRadius: BorderRadius.circular(8),
              ),
              child: SingleChildScrollView(
                child: Text(warningsText(), style: const TextStyle(fontSize: 12)),
              ),
            ),
          ],
          const SizedBox(height: 10),
          for (final item in processedItems.where((e) => !reviewOnly || e.needsReview))
            _swipeWrap(item, Card(
              margin: const EdgeInsets.symmetric(vertical: 4),
              child: ListTile(
                onTap: () => Navigator.push(
                  context,
                  MaterialPageRoute(
                    builder: (_) => ReceiptZoomPage(
                      path: item.path,
                      title: 'ظ…ط§ظƒظٹظ†ط© ${item.orderNum ?? 'طں'} - ${item.unitName ?? ''}',
                    ),
                  ),
                ),
                leading: Image.file(
                  File(item.path),
                  width: 48,
                  height: 48,
                  fit: BoxFit.cover,
                ),
                title: Text(
                  item.ignoredDoc || item.ign == 'drop'
                      ? 'ظ…ط³طھظ†ط¯ ظ…طھط¬ط§ظ‡ظ„ - ${baseName(item.path)}'
                      : 'ظ…ط§ظƒظٹظ†ط© ${item.orderNum ?? 'طں'} - ${item.unitName ?? 'ظˆط­ط¯ط© ط؛ظٹط± ظ…ط¹ط±ظˆظپط©'}',
                  style: const TextStyle(fontWeight: FontWeight.bold),
                ),
                subtitle: Text(subtitleOf(item)),
                trailing: Row(
                  mainAxisSize: MainAxisSize.min,
                  children: [
                    if (item.userExcluded)
                      IconButton(
                        tooltip: tr('ط±ط¬ظ‘ط¹ظ‡ (ط§ط­طھط³ط¨ظ‡)'),
                        icon: const Icon(Icons.undo, color: Colors.red),
                        onPressed: busy
                            ? null
                            : () async {
                                if (item.ign == 'drop') {
                                  keepIgnored(item);
                                  setState(() {});
                                } else {
                                  keepAsExtraBatch(item, processedItems);
                                  await rebuildFiles();
                                }
                              },
                      ),
                    if (item.needsReview)
                      Icon(Icons.warning_amber, color: _tintBorder(context, Colors.orange)),
                    IconButton(
                      icon: Icon(Icons.edit, color: Theme.of(context).colorScheme.primary),
                      onPressed: busy ? null : () => editReceipt(item),
                    ),
                  ],
                ),
              ),
            )),
        ],
      ),
    );
  }
}

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ ظ…طھطµظپط­ ط§ظ„ظپظˆظ„ط¯ط±ط§طھ â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
class FolderPicker extends StatefulWidget {
  final String start;
  const FolderPicker({super.key, required this.start});

  @override
  State<FolderPicker> createState() => _FolderPickerState();
}

class _FolderPickerState extends State<FolderPicker> {
  late String path = widget.start;
  List<Directory> dirs = [];
  String? error;

  @override
  void initState() {
    super.initState();
    load();
  }

  void load() {
    try {
      final l = Directory(path)
          .listSync()
          .whereType<Directory>()
          .where((d) => !baseName(d.path).startsWith('.'))
          .toList();
      l.sort((a, b) =>
          baseName(a.path).toLowerCase().compareTo(baseName(b.path).toLowerCase()));
      dirs = l;
      error = null;
    } catch (_) {
      dirs = [];
      error = 'ظ…ط´ ظ‚ط§ط¯ط± ط£ظپطھط­ ط§ظ„ظپظˆظ„ط¯ط± ط¯ظ‡.';
    }
  }

  void go(String p) => setState(() {
        path = p;
        load();
      });

  @override
  Widget build(BuildContext context) {
    final canUp = path.length > root.length;
    return Scaffold(
      appBar: AppBar(title: const Text('ط§ط®طھظٹط§ط± ظپظˆظ„ط¯ط±')),
      body: SafeArea(
        child: Column(
          children: [
            Padding(
              padding: const EdgeInsets.all(12),
              child: Text(path, textDirection: TextDirection.ltr),
            ),
            if (error != null)
              Padding(
                padding: const EdgeInsets.symmetric(horizontal: 12),
                child: Text(error!),
              ),
            Expanded(
              child: ListView(
                children: [
                  if (canUp)
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
                child: const Text('ط§ط®طھظٹط§ط± ظ‡ط°ط§ ط§ظ„ظپظˆظ„ط¯ط±'),
              ),
            ),
          ],
        ),
      ),
    );
  }
}


// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ ط´ط§ط´ط© ط§ظ„ظ…ط±ط§ط¬ط¹ط© â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// â”€â”€â”€â”€â”€â”€â”€â”€â”€ طµظˆط±ط© ط¨طھطھظƒط¨ظ‘ط± (ظ‚ط±طµط© ط¨طµط¨ط§ط¹ظٹظ† + ط¯ط¨ظ„ طھط§ط¨) ط¹ظ„ط´ط§ظ† طھط´ظˆظپ طھظپط§طµظٹظ„ ط§ظ„ط¨ط§طھط´ â”€â”€â”€â”€â”€â”€â”€â”€â”€
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
  final List<ReceiptItem>? allItems; // ظƒظ„ ط§ظ„ط¥ظٹطµط§ظ„ط§طھ (ظ„ط­ط³ط§ط¨ ط±ظ‚ظ… ط§ظ„ط¨ط§طھط´ ط§ظ„ط¥ط¶ط§ظپظٹ)
  final bool single; // ظپطھط­ظ†ط§ظ‡ط§ ظ…ظ† ط§ظ„ظ‚ظ„ظ… ظ„ط¥ظٹطµط§ظ„ ظˆط§ط­ط¯
  const ReviewPage({super.key, required this.items, this.allItems, this.single = false});

  @override
  State<ReviewPage> createState() => _ReviewPageState();
}

class _ReviewPageState extends State<ReviewPage> {
  int idx = 0;
  bool changed = false;
  int rev = 0; // ط¨ظٹطھط؛ظٹط± ظ„ظ…ط§ ط§ظ„طµظˆط±ط© طھطھط¯ظˆظ‘ط± ط¹ظ„ط´ط§ظ† طھطھط­ط¯ظ‘ط«
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
    if (v == null || v <= 0) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('ط§ظƒطھط¨ ظ…ط¨ظ„ط؛ طµط­ظٹط­.')),
      );
      return;
    }
    final it = item;
    it.amount = round2(v);
    it.batches = [it.amount];
    it.ok = true;
    it.alt = null;
    if (it.ignoredDoc) keepIgnored(it);
    it.how = 'manual';
    it.unusual = null;
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

  // طµظˆط±ط© ظ…طھط¬ط§ظ‡ظ„ط©: ط§ط¯ط®ظ„ظ‡ط§ ظƒط¨ط§طھط´طŒ ط£ظˆ طھط¬ط§ظ‡ظ„ظ‡ط§ ظ†ظ‡ط§ط¦ظٹ
  void resolveIgnored(bool keep) {
    final it = item;
    if (keep) {
      keepIgnored(it);
      changed = true;
      setState(() {});
      return;
    }
    it.ignoredDoc = false;
    it.userExcluded = true;
    it.ign = 'drop';
    saveFix(it);
    changed = true;
    done.add(it);
    if (widget.single) {
      Navigator.pop(context, true);
    } else if (idx < widget.items.length - 1) {
      go(idx + 1);
    } else {
      setState(() {});
    }
  }

  // طھط§ط±ظٹط® ظ…ط®طھظ„ظپ: ط³ظٹط¨ظ‡ ط¨طھط§ط±ظٹط®ظ‡طŒ ط£ظˆ ط§ظ†ظ‚ظ„ظ‡ ظ„طھط§ط±ظٹط® ط§ظ„ط£ط؛ظ„ط¨
  void resolveDate(bool toMain) {
    final it = item;
    if (toMain && it.dateMain != null) it.date = it.dateMain;
    it.dateFix = it.date;
    it.dateOdd = false;
    saveFix(it);
    changed = true;
    done.add(it);
    if (widget.single) {
      Navigator.pop(context, true);
    } else if (idx < widget.items.length - 1) {
      go(idx + 1);
    } else {
      setState(() {});
    }
  }

  // ظ‚ط±ط§ط± ط§ظ„ط¥ظٹطµط§ظ„ ط§ظ„ظ…ط´ظƒظˆظƒ ط¥ظ†ظ‡ ظ…ظƒط±ط±: ط§ط³طھط¨ط¹ط¯ظ‡طŒ ط£ظˆ ط§ط­طھط³ط¨ظ‡ ظƒط¨ط§طھط´ ط¥ط¶ط§ظپظٹ
  void resolveDup(bool keep) {
    final it = item;
    final all = widget.allItems ?? widget.items;
    if (keep) {
      keepAsExtraBatch(it, all);
    } else {
      it.dupSuspect = false;
      it.userExcluded = true;
      it.dupDecision = 'drop';
      saveFix(it);
    }
    changed = true;
    done.add(it);
    if (widget.single) {
      Navigator.pop(context, true);
    } else if (idx < widget.items.length - 1) {
      go(idx + 1);
    } else {
      setState(() {});
    }
  }

  Future<void> rotate() async {
    try {
      await fixImageAsync(item.path, item.path, 90);
      await FileImage(File(item.path)).evict();
      imageCache.clear();
      changed = true;
      setState(() => rev++);
    } catch (_) {
      if (!mounted) return;
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('ظ…ظ‚ط¯ط±طھط´ ط£ط¯ظˆظ‘ط± ط§ظ„طµظˆط±ط©.')),
      );
    }
  }

  @override
  Widget build(BuildContext context) {
    AppSettings.of(context);
    final it = item;
    final image = Container(
      color: Colors.black12,
      child: Stack(
        children: [
          Positioned.fill(
            child: ZoomableImage(key: ValueKey('${it.path}#$rev'), path: it.path),
          ),
          Positioned(
            top: 4,
            left: 4,
            child: IconButton.filledTonal(
              tooltip: tr('ظ…ظ„ط، ط§ظ„ط´ط§ط´ط©'),
              icon: const Icon(Icons.fullscreen),
              onPressed: () => Navigator.push(
                context,
                MaterialPageRoute(
                  builder: (_) => ReceiptZoomPage(
                    path: it.path,
                    title: 'ظ…ط§ظƒظٹظ†ط© ${it.orderNum ?? 'طں'} - ${it.unitName ?? ''}',
                  ),
                ),
              ),
            ),
          ),
        ],
      ),
    );

    final form = SingleChildScrollView(
      padding: const EdgeInsets.all(12),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.stretch,
        children: [
          if (it.dupSuspect || it.userExcluded)
            Container(
              margin: const EdgeInsets.only(bottom: 8),
              padding: const EdgeInsets.all(8),
              decoration: BoxDecoration(
                color: _tintBg(context, Colors.orange, soft: true),
                border: Border.all(color: _tintBorder(context, Colors.orange)),
                borderRadius: BorderRadius.circular(8),
              ),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.stretch,
                children: [
                  Text(
                    it.userExcluded
                        ? 'ًںڑ« ط§ظ„ط¥ظٹطµط§ظ„ ط¯ظ‡ ظ…ط³طھط¨ط¹ط¯ (ظ…ظƒط±ط±).'
                        : 'âڑ ï¸ڈ ظ…ط´ظƒظˆظƒ ط¥ظ†ظ‡ ظ…ظƒط±ط±: ظ†ظپط³ ط§ظ„ظ…ط§ظƒظٹظ†ط© ${it.orderNum ?? ''} ظˆظ†ظپط³ ط§ظ„ظ…ط¨ظ„ط؛. '
                            'ظ…ط´ ظ…ط­ط³ظˆط¨ ظپظٹ ط§ظ„ط¥ط¬ظ…ط§ظ„ظٹ ظ„ط­ط¯ ظ…ط§ طھظ‚ط±ط±. ط²ظˆظ‘ظ… ط¹ظ„ظ‰ ط§ظ„ط¨ط§طھط´ ظˆظ‚ط§ط±ظ†.',
                    style: TextStyle(color: _tintFg(context, Colors.orange), fontSize: 13),
                  ),
                  const SizedBox(height: 6),
                  Row(
                    children: [
                      Expanded(
                        child: FilledButton(
                          onPressed: () => resolveDup(true),
                          child: const Text('ظ…ط´ ظ…ظƒط±ط± - ط§ط­طھط³ط¨ظ‡ ط¨ط§طھط´ ط¥ط¶ط§ظپظٹ',
                              textAlign: TextAlign.center, style: TextStyle(fontSize: 12)),
                        ),
                      ),
                      if (!it.userExcluded) ...[
                        const SizedBox(width: 6),
                        Expanded(
                          child: OutlinedButton(
                            onPressed: () => resolveDup(false),
                            child: const Text('ظ…ظƒط±ط± - ط§ط³طھط¨ط¹ط¯ظ‡',
                                textAlign: TextAlign.center, style: TextStyle(fontSize: 12)),
                          ),
                        ),
                      ],
                    ],
                  ),
                ],
              ),
            ),
          if (it.ignoredDoc)
            Container(
              margin: const EdgeInsets.only(bottom: 8),
              padding: const EdgeInsets.all(8),
              decoration: BoxDecoration(
                color: _tintBg(context, Colors.orange, soft: true),
                border: Border.all(color: _tintBorder(context, Colors.orange)),
                borderRadius: BorderRadius.circular(8),
              ),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.stretch,
                children: [
                  Text(
                    'ًں“„ ط§ظ„طھط·ط¨ظٹظ‚ ط§طھط¬ط§ظ‡ظ„ ط§ظ„طµظˆط±ط© ط¯ظٹ ظ„ط£ظ† ط´ظƒظ„ظ‡ط§ ط¥ظٹطµط§ظ„ ط¯ظپط¹ ط£ظˆ ظ…ط³طھظ†ط¯ ظ…ط´ ط¨ط§طھط´. '
                    'ط²ظˆظ‘ظ… ظˆط´ظˆظپ: ظ„ظˆ ط¨ط§طھط´ ط­ظ‚ظٹظ‚ظٹ ط§ط¯ط®ظ„ظ‡ ظˆظƒظ…ظ‘ظ„ ط§ظ„ظ…ط¨ظ„ط؛ ظˆط§ظ„ظ…ط§ظƒظٹظ†ط©.',
                    style: TextStyle(color: _tintFg(context, Colors.orange), fontSize: 13),
                  ),
                  const SizedBox(height: 6),
                  Row(
                    children: [
                      Expanded(
                        child: FilledButton(
                          onPressed: () => resolveIgnored(true),
                          child: const Text('ط¯ظ‡ ط¨ط§طھط´ - ط§ط¯ط®ظ„ظ‡',
                              textAlign: TextAlign.center, style: TextStyle(fontSize: 12)),
                        ),
                      ),
                      const SizedBox(width: 6),
                      Expanded(
                        child: OutlinedButton(
                          onPressed: () => resolveIgnored(false),
                          child: const Text('ظ…ط´ ط¨ط§طھط´ - طھط¬ط§ظ‡ظ„ظ‡',
                              textAlign: TextAlign.center, style: TextStyle(fontSize: 12)),
                        ),
                      ),
                    ],
                  ),
                ],
              ),
            ),
          if (it.dateOdd)
            Container(
              margin: const EdgeInsets.only(bottom: 8),
              padding: const EdgeInsets.all(8),
              decoration: BoxDecoration(
                color: _tintBg(context, Colors.orange, soft: true),
                border: Border.all(color: _tintBorder(context, Colors.orange)),
                borderRadius: BorderRadius.circular(8),
              ),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.stretch,
                children: [
                  Text(
                    'ًں“… طھط§ط±ظٹط® ط§ظ„ط¥ظٹطµط§ظ„ ${it.date} ط¨ظٹظ†ظ…ط§ ط£ط؛ظ„ط¨ ط§ظ„ط¥ظٹطµط§ظ„ط§طھ ${it.dateMain}.',
                    style: TextStyle(color: _tintFg(context, Colors.orange), fontSize: 13),
                  ),
                  const SizedBox(height: 6),
                  Row(
                    children: [
                      Expanded(
                        child: FilledButton(
                          onPressed: () => resolveDate(true),
                          child: Text('ط®ظ„ظٹظ‡ ط¨طھط§ط±ظٹط® ${it.dateMain}',
                              textAlign: TextAlign.center, style: const TextStyle(fontSize: 12)),
                        ),
                      ),
                      const SizedBox(width: 6),
                      Expanded(
                        child: OutlinedButton(
                          onPressed: () => resolveDate(false),
                          child: const Text('ط§ظ„طھط§ط±ظٹط® طµط­ - ط³ظٹط¨ظ‡',
                              textAlign: TextAlign.center, style: TextStyle(fontSize: 12)),
                        ),
                      ),
                    ],
                  ),
                ],
              ),
            ),
          if (it.weakId && !it.ignoredDoc && !it.dupSuspect && !it.userExcluded && it.row != null)
            Padding(
              padding: const EdgeInsets.only(bottom: 8),
              child: Text(
                it.how == 'auto-number' || it.how == 'conflict'
                    ? 'â‌“ ط±ظ‚ظ… ط§ظ„ظ…ط§ظƒظٹظ†ط© ط§طھط­ط¯ط¯ طھظ„ظ‚ط§ط¦ظٹط§ظ‹ (${it.orderNum}) - ط§طھط£ظƒط¯ ظ…ظ†ظ‡ ظˆط§ط¶ط؛ط· ط­ظپط¸.'
                    : 'â‌“ ط§ظ„ظˆط­ط¯ط© ط§طھط­ط¯ط¯طھ ط¨ط§ظ„ط§ط³ظ…/ط§ظ„ظ…ظٹط±ط´ط§ظ†طھ ظپظ‚ط· - ط§طھط£ظƒط¯ ظ…ظ† ط§ظ„ظ…ط§ظƒظٹظ†ط© ظˆط§ط¶ط؛ط· ط­ظپط¸.',
                style: TextStyle(color: _tintFg(context, Colors.orange), fontSize: 13),
              ),
            ),
          if (it.unusual != null)
            Padding(
              padding: const EdgeInsets.only(bottom: 8),
              child: Text('âڑ ï¸ڈ ${it.unusual}  (ظ„ظˆ ط§ظ„ظ…ط¨ظ„ط؛ طµط­ ط§ط¶ط؛ط· ط­ظپط¸)',
                  style: TextStyle(color: _tintFg(context, Colors.orange), fontSize: 13)),
            ),
          Text(
            'ط§ظ„ظ…ظ‚ط±ظˆط،: ${it.amount > 0 ? it.amount.toStringAsFixed(2) : 'ط؛ظٹط± ظ…ظ‚ط±ظˆط،'} ط¬.ظ…'
            '${it.batches.length > 1 ? '  (= ${it.batches.map((b) => b.toStringAsFixed(2)).join(' + ')})' : ''}',
            style: const TextStyle(fontWeight: FontWeight.bold),
          ),
          if (it.alt != null)
            Padding(
              padding: const EdgeInsets.only(top: 4),
              child: Row(
                children: [
                  Expanded(
                    child: Text('âڑ ï¸ڈ ط§ظ„ظ‚ط±ط§ط،ط© ط§ظ„طھط§ظ†ظٹط©: ${it.alt!.toStringAsFixed(2)}',
                        style: TextStyle(color: _tintFg(context, Colors.orange))),
                  ),
                  TextButton(
                    onPressed: () => setState(() => amt.text = it.alt!.toStringAsFixed(2)),
                    child: const Text('ط§ط³طھط®ط¯ظ…ظ‡ط§'),
                  ),
                ],
              ),
            ),
          const SizedBox(height: 8),
          TextField(
            controller: amt,
            keyboardType: const TextInputType.numberWithOptions(decimal: true),
            decoration: InputDecoration(
              labelText: tr('ط§ظ„ظ…ط¨ظ„ط؛ ط§ظ„ط¥ط¬ظ…ط§ظ„ظٹ'),
              border: const OutlineInputBorder(),
            ),
          ),
          const SizedBox(height: 8),
          DropdownButtonFormField<int>(
            value: order,
            isExpanded: true,
            hint: const Text('ط§ط®طھط± ط§ظ„ظ…ط§ظƒظٹظ†ط©'),
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
          Row(
            children: [
              OutlinedButton.icon(
                onPressed: rotate,
                icon: const Icon(Icons.rotate_right),
                label: const Text('طھط¯ظˆظٹط±'),
              ),
              const SizedBox(width: 8),
              Expanded(
                child: FilledButton(
                  onPressed: () => save(),
                  child: Text(idx < widget.items.length - 1 ? 'ط­ظپط¸ ظˆط§ظ„طھط§ظ„ظٹ' : 'ط­ظپط¸'),
                ),
              ),
            ],
          ),
        ],
      ),
    );

    return PopScope(
      canPop: false,
      onPopInvokedWithResult: (didPop, _) {
        if (!didPop) Navigator.pop(context, changed);
      },
      child: Scaffold(
        appBar: AppBar(
          title: Text('ظ…ط±ط§ط¬ط¹ط© ${idx + 1} / ${widget.items.length}'
              '${done.contains(it) ? '  âœ“' : ''}'),
          leading: IconButton(
            icon: const Icon(Icons.arrow_back),
            onPressed: () => Navigator.pop(context, changed),
          ),
          actions: [
            IconButton(
              icon: const Icon(Icons.navigate_before),
              tooltip: tr('ط§ظ„ط³ط§ط¨ظ‚'),
              onPressed: idx > 0 ? () => go(idx - 1) : null,
            ),
            IconButton(
              icon: const Icon(Icons.navigate_next),
              tooltip: tr('ط§ظ„طھط§ظ„ظٹ'),
              onPressed: idx < widget.items.length - 1 ? () => go(idx + 1) : null,
            ),
          ],
        ),
        body: SafeArea(
          child: LayoutBuilder(
            builder: (context, c) {
              if (c.maxWidth > 600) {
                return Row(
                  children: [
                    Expanded(flex: 3, child: image),
                    Expanded(flex: 2, child: form),
                  ],
                );
              }
              return Column(
                children: [
                  Expanded(flex: 5, child: image),
                  Expanded(flex: 3, child: form),
                ],
              );
            },
          ),
        ),
      ),
    );
  }
}


// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ ظ…ظ„ظپط§طھ ط§ظ„ظ…ط§ظƒظٹظ†ط§طھ (PDF ظ„ظƒظ„ ظ…ط§ظƒظٹظ†ط©) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
String? loadPdfDir() => _readCfg()['pdfDir'] as String?;

// ظپظˆظ„ط¯ط±ط§طھ ظ…ط³طھظ†ط¯ط§طھ ظˆط§طھط³ط§ط¨ (ط¬ظ†ط¨ ظپظˆظ„ط¯ط±ط§طھ ط§ظ„طµظˆط±)
List<Directory> findWhatsAppDocDirs() {
  final out = <Directory>[];
  for (final d in findWhatsAppImageDirs()) {
    final p = d.path;
    if (p.endsWith('WhatsApp Images')) {
      final dd = Directory('${p.substring(0, p.length - 'WhatsApp Images'.length)}WhatsApp Documents');
      for (final sub in ['', '/Sent', '/Private']) {
        final x = Directory('${dd.path}$sub');
        if (x.existsSync() && !out.any((o) => o.path == x.path)) out.add(x);
      }
    }
  }
  return out;
}

Set<String> loadDoneM() {
  final l = _readCfg()['doneM'];
  return l is List ? l.map((e) => e.toString()).toSet() : <String>{};
}

void saveDoneM(Set<String> s) {
  final m = _readCfg();
  final l = s.toList();
  m['doneM'] = l.length > 500 ? l.sublist(l.length - 500) : l;
  _writeCfg(m);
}

void savePdfDir(String? f) {
  final m = _readCfg();
  m['pdfDir'] = f;
  _writeCfg(m);
}

// ط§ظ„ط§ط³ظ… ط§ظ„ط¬ط¯ظٹط¯: "1.ط§ط³ظ….ط§ظ„ظˆط­ط¯ط©.1.ظٹظˆظ….4.11" (ظˆط­ط¯ط© ظپظٹظ‡ط§ ط£ظƒطھط± ظ…ظ† ظ…ط§ظƒظٹظ†ط©) ط£ظˆ "17.ط§ط³ظ….ط§ظ„ظˆط­ط¯ط©.ظٹظˆظ….4.11"
String machineNewName(TerminalInfo info, String day, String month) {
  final same = terminals.values.where((t) => t.excelRow == info.excelRow).toList()
    ..sort((a, b) => a.orderNum.compareTo(b.orderNum));
  final unit = info.unitName.trim().replaceAll(RegExp(r'\s+'), '.');
  final k = same.indexWhere((t) => t.orderNum == info.orderNum) + 1;
  return [
    info.orderNum.toString(),
    unit,
    if (same.length > 1) '$k',
    'ظٹظˆظ…',
    day,
    month,
  ].join('.');
}

// â”€â”€â”€â”€â”€â”€â”€â”€â”€ ط¹ط¯ظ‘ ط¥ظٹطµط§ظ„ط§طھ ط§ظ„ط¯ظپط¹: ظƒظ„ ط¥ظٹطµط§ظ„ ط¯ظپط¹ ظ„ظٹظ‡ ط³ط·ط± RRN (12 ط±ظ‚ظ…طŒ ط¨ظٹط¨ط¯ط£ ط¨ظ€ 63) â”€â”€â”€â”€â”€â”€â”€â”€â”€
// ط¨ظ†ظ‚ط±ط§ ط§ظ„طµظپط­ط© ط¹ظ„ظ‰ ط´ط±ط§ط¦ط­ ط±ط£ط³ظٹط© ط¶ظٹظ‚ط© (ط§ظ„ظ‚ط±ط§ط،ط© ط£ط¯ظ‚ ظ…ظ† ط§ظ„طµظپط­ط© ظƒظ„ظ‡ط§)طŒ ظˆط¨ظ†ط¬ظ…ط¹ "ظ†ظ‚ط§ط· ط§ط±طھظƒط§ط²"
// (ط³ط·ط± RRN) ظ„ظƒظ„ ط¥ظٹطµط§ظ„. ط¥ظٹطµط§ظ„ط§طھ ط§ظ„ط¥ظ‚ظپط§ظ„ ظˆط§ظ„ظ…ظ„ط®طµ ظ…ط§ ظپظٹظ‡ط§ط´ RRN ظپظ…ط¨طھطھط¹ط¯ظ‘ط´.
final RegExp _dateRe = RegExp(r'(20\d{2})[^\d\n]{1,3}(\d{2})[^\d\n]{1,3}(\d{2})');

String? _dayKey(String text) {
  for (final m in _dateRe.allMatches(text)) {
    final mo = int.parse(m.group(2)!), d = int.parse(m.group(3)!);
    if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31) return '$d.$mo';
  }
  return null;
}

// null = ظ…ظپظٹط´ RRN ظپظٹ ط§ظ„ط³ط·ط± | 12 ط±ظ‚ظ… = RRN ظƒط§ظ…ظ„ | 'P' = RRN ظ…ظ‚ط±ظˆط، ظ†ط§ظ‚طµ
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

// ط¨ظٹظ‚طµ ط§ظ„طµظپط­ط© ظ„ظ€ 5 ط´ط±ط§ط¦ط­ ط±ط£ط³ظٹط© (ط£ط¹ظ…ط¯ط©) ظ…طھط¯ط§ط®ظ„ط© ظˆظٹظƒط¨ظ‘ط±ظ‡ط§. ط¨ظٹط±ط¬ظ‘ط¹ [ط¹ط±ط¶طŒ ط§ط±طھظپط§ط¹طŒ x0 ظ„ظƒظ„ ط´ط±ظٹط­ط©...]
List<int> _makeStrips(String src, String prefix) {
  final im = img.decodeImage(File(src).readAsBytesSync());
  if (im == null) return [];
  final w = im.width, h = im.height;
  const n = 5;
  final cw = w / n;
  final out = <int>[w, h];
  for (var i = 0; i < n; i++) {
    final x0 = math.max(0, (i * cw - cw * 0.12).round());
    final x1 = math.min(w, ((i + 1) * cw + cw * 0.12).round());
    var c = img.copyCrop(im, x: x0, y: 0, width: x1 - x0, height: h);
    c = img.copyResize(c,
        width: (c.width * _stripScale).round(), interpolation: img.Interpolation.linear);
    File('${prefix}_$i.jpg').writeAsBytesSync(img.encodeJpg(c, quality: 92));
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
  final StringBuffer text = StringBuffer();
  double pageH = 3000;

  void add(Anchor n) {
    final list = pages.last;
    for (final a in list) {
      // ظ†ظپط³ ط§ظ„ط¥ظٹطµط§ظ„ ط§طھظ‚ط±ط§ ظ…ظ† ط´ط±ظٹط­طھظٹظ† ظ…طھط¯ط§ط®ظ„طھظٹظ†: ظ†ظپط³ ط§ظ„ظ…ظƒط§ظ† ط¹ظ„ظ‰ ط§ظ„طµظپط­ط©
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

    // ط§ظ„ظ…ط§ظƒظٹظ†ط©: ط£ظƒطھط± Terminal ID ظ…طھظƒط±ط±
    final tc = <String, int>{};
    for (final m in RegExp(r'(?<!\d)(\d{8})(?!\d)').allMatches(all)) {
      var c = m.group(1)!;
      c = tidFix[c] ?? c;
      if (terminals.containsKey(c)) tc[c] = (tc[c] ?? 0) + 1;
    }
    r.tid = tc.isNotEmpty
        ? tc.entries.reduce((a, b) => a.value >= b.value ? a : b).key
        : (findTerminalIn([all]) ?? discoverTerminal([all]));

    // ط§ظ„ظٹظˆظ… ط§ظ„ط£ط؛ظ„ط¨ + ط§ظ„ط¥ظٹطµط§ظ„ط§طھ ط§ظ„ظ„ظٹ ظ…ط´ طھط¨ط¹ ظ†ظپط³ ط§ظ„ظٹظˆظ…
    if (days.isEmpty) {
      for (final m in _dateRe.allMatches(all)) {
        final k = _dayKey(m.group(0)!);
        if (k != null) days[k] = (days[k] ?? 0) + 1;
      }
    }
    if (days.isNotEmpty) {
      final best = days.entries.reduce((a, b) => a.value >= b.value ? a : b).key;
      final p = best.split('.');
      r.day = p[0];
      r.month = p[1];
      r.rrnDay.forEach((rrn, d) {
        if (d != null && d != best) r.otherDay.add('ط¥ظٹطµط§ظ„ ${rrn.substring(6)} ط¨طھط§ط±ظٹط® $d');
      });
      for (final d in partialDays) {
        if (d != best) r.otherDay.add('ط¥ظٹطµط§ظ„ (ط±ظ‚ظ…ظ‡ ظ…ط´ ظˆط§ط¶ط­) ط¨طھط§ط±ظٹط® $d');
      }
    }

    // طھط³ظ„ط³ظ„ ط£ط±ظ‚ط§ظ… ط§ظ„ط¥ظٹطµط§ظ„ط§طھ (ط¢ط®ط± 6 ط£ط±ظ‚ط§ظ… ظ…ظ† ط§ظ„ظ€ RRN): ط£ظƒط¨ط± ظ…ط¬ظ…ظˆط¹ط© ظ…طھطھط§ظ„ظٹط©
    final nums = r.rrnDay.keys.map((k) => int.parse(k.substring(6))).toList()..sort();
    if (nums.length >= 2) {
      var bestStart = 0, bestLen = 1, start = 0;
      for (var i = 1; i <= nums.length; i++) {
        if (i == nums.length || nums[i] - nums[i - 1] > 15) {
          if (i - start > bestLen) {
            bestLen = i - start;
            bestStart = start;
          }
          start = i;
        }
      }
      final cl = nums.sublist(bestStart, bestStart + bestLen);
      r.seqMin = cl.first;
      r.seqMax = cl.last;
      final have = cl.toSet();
      for (var n = cl.first; n <= cl.last; n++) {
        if (!have.contains(n)) r.seqMissing.add(n);
      }
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
  bool get agree => rangeLen == null || count == rangeLen;
  String get rangeText =>
      seqMin == null ? '' : ' â€¢ ط§ظ„طھط³ظ„ط³ظ„ ${seqMin.toString().padLeft(6, '0')}-${seqMax.toString().padLeft(6, '0')} ($rangeLen)';
  String get seqNote {
    final miss = seqMissing.take(10).map((n) => n.toString().padLeft(6, '0')).join('طŒ ');
    return 'ط§ظ„ط¹ط¯ط¯ $count ط¨ط³ ط§ظ„طھط³ظ„ط³ظ„ $rangeLen'
        '${seqMissing.isNotEmpty ? ' - ط£ط±ظ‚ط§ظ… ظ…ط´ ظ„ط§ظ‚ظٹظ‡ط§: $miss${seqMissing.length > 10 ? ' ...' : ''}' : ''}'
        ' (ظ…ظ…ظƒظ† ط¥ظٹطµط§ظ„ط§طھ ظ†ط§ظ‚طµط© ط£ظˆ ظ…ط´ ط§طھظ‚ط±طھطŒ ط±ط§ط¬ط¹ ط§ظ„ظ…ظ„ظپ)';
  }
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

// ظ‚ط±ط§ط،ط© ط§ظ„طµظپط­ط© ظƒظ„ظ‡ط§ + ظ‚ط±ط§ط،ط© ط´ط±ط§ط¦ط­ ظ…ظƒط¨ظ‘ط±ط©طŒ ظˆط¨ط¹ط¯ظٹظ† ط¯ظ…ط¬ ط§ظ„ظ†طھط§ظٹط¬ (ظ†ظپط³ ط§ظ„ط¥ظٹطµط§ظ„ ط¨ظٹطھط¹ط¯ظ‘ ظ…ط±ط©)
int _tmStrips = 0, _tmOcr = 0; // طھط´ط®ظٹطµ: ظ…ط¬ظ…ظˆط¹ ط§ظ„ظˆظ‚طھ (ms) ظ„ط¹ظ…ظ„ ط§ظ„ط´ط±ط§ط¦ط­ ظˆظ„ظ„ظ€ OCR

Future<void> ocrPageStrips(
    TextRecognizer rec, String pngPath, String prefix, MCollector col) async {
  final sw = Stopwatch()..start();
  // طھظ‚ط·ظٹط¹ ط§ظ„ط´ط±ط§ط¦ط­ (Dart طµط±ظپ ظˆط¨ط·ظٹط،) ط¨ظٹط´طھط؛ظ„ ظپظٹ isolate ط¨ط§ظ„طھظˆط§ط²ظٹ ظ…ط¹ OCR ط§ظ„طµظپط­ط© ط§ظ„ظƒط§ظ…ظ„ط©
  final stripsF = Isolate.run(() => _makeStrips(pngPath, prefix))
      .then<List<int>>((v) => v, onError: (_) => <int>[]);
  RecognizedText? whole;
  try {
    whole = await rec.processImage(InputImage.fromFilePath(pngPath));
  } catch (_) {}
  final tWhole = sw.elapsedMilliseconds;
  final info = await stripsF;
  _tmStrips += sw.elapsedMilliseconds - tWhole;
  col.pageH = info.length >= 2 ? info[1].toDouble() : 3000;
  col.pages.add(<Anchor>[]);
  if (whole != null) _collect(whole, 0, 1, col);
  final t1 = sw.elapsedMilliseconds;
  for (var i = 2; i < info.length; i++) {
    final fp = '${prefix}_${i - 2}.jpg';
    try {
      final res = await rec.processImage(InputImage.fromFilePath(fp));
      _collect(res, info[i].toDouble(), _stripScale, col);
    } catch (_) {}
    try {
      File(fp).deleteSync();
    } catch (_) {}
  }
  _tmOcr += tWhole + (sw.elapsedMilliseconds - t1);
}

// â”€â”€â”€â”€â”€â”€â”€â”€â”€ ط­ط§ظ„ط© ظپط­طµ ط¥ظٹطµط§ظ„ط§طھ ط§ظ„ط¯ظپط¹: ط¨طھطھط­ظپط¸ ط¨ط¹ط¯ ظƒظ„ ظ…ظ„ظپ ط¹ظ„ط´ط§ظ† ظ„ظˆ ط§ظ„طھط·ط¨ظٹظ‚ ط§طھظ‚ظپظ„ ظٹظƒظ…ظ‘ظ„ ظ…ظ† ظ…ظƒط§ظ†ظ‡ â”€â”€â”€â”€â”€â”€â”€â”€â”€
final String mStatePath = '$baseDir/.app_machines.json';
final String defaultMDest = '$baseDir/ظˆط­ط¯ط§طھ';

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
  int count = -1; // -1 = ظ„ط³ظ‡ ظ…ط§ ط§طھط¹ط¯ظ‘ط´
  int tries = 0;
  String detail = '';
  final List<String> idNotes = [];
  final List<String> cntNotes = [];
  List<String> altTids = []; // ظ…ط±ط´ط­ظٹظ† طھط§ظ†ظٹظٹظ† ظ„ظ„ظ…ط§ظƒظٹظ†ط© (ظ…ط´ ط¨ظٹطھط­ظپط¸ظˆط§)
  Map<String, double> votes = {};
  String? tidUsed, dayUsed, monthUsed;
  bool idOk = true;
  bool cntOk = true;
  int verdict = 0; // 0 = ظ„ط³ظ‡ ظ…ط§ ط§طھط±ط§ط¬ط¹ط´طŒ 1 = طµط­ (ظٹظپط¶ظ„ ظ…ط¹ ظ…ظ„ظپط§طھ ط§ظ„ظٹظˆظ…)طŒ -1 = ط؛ظ„ط· (ظ…ط³طھط¨ط¹ط¯)
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
    return l
        .whereType<Map>()
        .map((e) => MResult.fromJson(Map<String, dynamic>.from(e)))
        .toList();
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

// طھط¹ط±ظٹظپ ط³ط±ظٹط¹: ط§ظ„ظ…ط§ظƒظٹظ†ط© (Terminal ID) + ط§ظ„ظٹظˆظ… ظ…ظ† ظ†طµ طµظپط­ط© ط£ظˆ ط£ظƒطھط±
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

// â”€â”€â”€â”€â”€â”€â”€â”€â”€ طھط¹ط±ظٹظپ ط§ظ„ظ…ط§ظƒظٹظ†ط© ظ…ظ† ط¬ظˆظ‡ ط§ظ„ظ€ PDF ط¨ط³ (طھطµظˆظٹطھ ط¨ظٹظ† ط§ظ„طµظپط­ط§طھ) â€” ظ…ظپظٹط´ ط§ط¹طھظ…ط§ط¯ ط¹ظ„ظ‰ ط§ط³ظ… ط§ظ„ظ…ظ„ظپ â”€â”€â”€â”€â”€â”€â”€â”€â”€
// ط£طµظˆط§طھ طµظپط­ط© ظˆط§ط­ط¯ط©: "TERMINAL ID" ط¨ظˆط²ظ† 3طŒ ط£ظٹ ط±ظ‚ظ… 8 ط®ط§ظ†ط§طھ ظ…ط¹ط±ظˆظپ ط¨ظˆط²ظ† 1طŒ ظˆطھظ‚ط±ظٹط¨ظٹ (ط؛ظ„ط·ط© ط±ظ‚ظ…) ط¨ظˆط²ظ† 0.5
Map<String, double> _tidVotes(String text) {
  final v = <String, double>{};
  final lab = _tidLabeled(text);
  if (lab != null) v[lab] = (v[lab] ?? 0) + 3;
  for (final m in RegExp(r'(?<!\d)(\d{8})(?!\d)').allMatches(text)) {
    final c0 = m.group(1)!;
    final c = tidFix[c0] ?? c0;
    if (terminals.containsKey(c)) v[c] = (v[c] ?? 0) + 1;
  }
  if (v.isEmpty) {
    final f = findTerminalFuzzyIn([text], null);
    if (f != null) v[f] = 0.5;
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

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ طھظ‚ط±ظٹط± ط¥ظٹطµط§ظ„ط§طھ ط§ظ„ط¯ظپط¹ â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
String _l(BuildContext c, String ar, String en) => AppSettings.of(c).isAr ? ar : en;

class _NoteKind {
  final String key;
  final int sev; // 2 = ظ„ط§ط²ظ… طھط±ط§ط¬ط¹طŒ 1 = ط§ظ†طھط¨ظ‡طŒ 0 = ظ…ط¹ظ„ظˆظ…ط©
  final IconData icon;
  final String titleAr, titleEn, whyAr, whyEn, fixAr, fixEn;
  const _NoteKind(this.key, this.sev, this.icon, this.titleAr, this.titleEn, this.whyAr, this.whyEn,
      this.fixAr, this.fixEn);
}

const List<_NoteKind> _kNoteKinds = [
  _NoteKind(
      'seq',
      2,
      Icons.format_list_numbered,
      'ط§ظ„ط¹ط¯ط¯ ظ…ط´ ظ…ط·ط§ط¨ظ‚ ظ„طھط³ظ„ط³ظ„ ط§ظ„ط¥ظٹطµط§ظ„ط§طھ',
      'Count does not match the receipt sequence',
      'ط£ط±ظ‚ط§ظ… ط§ظ„ط¥ظٹطµط§ظ„ط§طھ (RRN) ظ…طھطھط§ظ„ظٹط©طŒ ظپط§ظ„ط¹ط¯ط¯ ط§ظ„ظ…طھظˆظ‚ط¹ ظ‡ظˆ ظ…ظ† ط£ظˆظ„ ط±ظ‚ظ… ظ„ط¢ط®ط± ط±ظ‚ظ…. ط§ظ„ط¹ط¯ط¯ ط§ظ„ظ„ظٹ ط§طھط¹ط¯ظ‘ ظ…ط®طھظ„ظپ ط¹ظ†ظ‡طŒ ظٹط¹ظ†ظٹ ط؛ط§ظ„ط¨ط§ظ‹ ظپظٹظ‡ ط¥ظٹطµط§ظ„ ظ†ط§ظ‚طµ ظپظٹ ط§ظ„ظ…ظ„ظپ ط£ظˆ ط§ظ„طھط·ط¨ظٹظ‚ ظ…ط§ ظ‚ط±ط§ط´ ط¥ظٹطµط§ظ„.',
      'Receipt numbers (RRN) are consecutive, so the expected count is first-to-last number. The counted number differs, so a receipt is probably missing from the file or was not read.',
      'ط§ظپطھط­ ط§ظ„ظ…ظ„ظپ ظˆط±ط§ط¬ط¹ ط§ظ„ط£ط±ظ‚ط§ظ… ط§ظ„ظ„ظٹ ظ…ط´ ظ„ط§ظ‚ظٹظ‡ط§ (ظ…ظƒطھظˆط¨ط© طھط­طھ)طŒ ظˆط³ط¬ظ‘ظ„ ط§ظ„ط¹ط¯ط¯ ط§ظ„طµط­ ظپظٹ ط®ط§ظ†ط© "ط§ظ„طµط­" ظ„ظˆ ط§طھط£ظƒط¯طھ.',
      'Open the file and check the missing numbers listed below, then enter the right count in the "Correct" field if you are sure.'),
  _NoteKind(
      'otherday',
      2,
      Icons.event_busy,
      'ظپظٹظ‡ ط¥ظٹطµط§ظ„ط§طھ ظ…ظ† ظٹظˆظ… طھط§ظ†ظٹ',
      'Receipts from another day',
      'طھط§ط±ظٹط® ط¨ط¹ط¶ ط§ظ„ط¥ظٹطµط§ظ„ط§طھ ظ…ط®طھظ„ظپ ط¹ظ† ظٹظˆظ… ط¨ط§ظ‚ظٹ ط§ظ„ظ…ظ„ظپطŒ ظپظ…ظ…ظƒظ† طھظƒظˆظ† ط§طھط¶ط§ظپطھ ط¨ط§ظ„ط؛ظ„ط· ط£ظˆ ط§ظ„ظ…ظ„ظپ ظپظٹظ‡ ط£ظƒطھط± ظ…ظ† ظٹظˆظ….',
      'Some receipts have a different date from the rest of the file; they may have been added by mistake or the file spans more than one day.',
      'ط±ط§ط¬ط¹ ط§ظ„طھظˆط§ط±ظٹط® ط§ظ„ظ…ط°ظƒظˆط±ط©طŒ ظˆظ„ظˆ ط§ظ„ظ…ظ„ظپ ظ…ط´ ط¨طھط§ط¹ ط§ظ„ظٹظˆظ… ط¹ظ„ظ‘ظ…ظ‡ ط؛ظ„ط· âœ—.',
      'Check the dates listed; if the file is not for this day, mark it wrong âœ—.'),
  _NoteKind(
      'machconf',
      2,
      Icons.compare_arrows,
      'طھط¹ط§ط±ط¶ ظپظٹ ط±ظ‚ظ… ط§ظ„ظ…ط§ظƒظٹظ†ط©',
      'Machine mismatch',
      'ط§ط³ظ… ط§ظ„ظ…ظ„ظپ ط§طھط­ط¯ط¯ ظ„ظ…ط§ظƒظٹظ†ط©طŒ ظ„ظƒظ† ط§ظ„ط¥ظٹطµط§ظ„ط§طھ ط¬ظˆط§ظ‡ ط¨طھظ‚ظˆظ„ ط¥ظ†ظ‡ط§ ظ…ط§ظƒظٹظ†ط© طھط§ظ†ظٹط©.',
      'The file was named for one machine, but the receipts inside belong to another machine.',
      'ط§ط³طھط®ط¯ظ… "طھط­ط¯ظٹط¯ ط§ظ„ظ…ط§ظƒظٹظ†ط©" ظ…ظ† ط§ظ„ظ‚ط§ظٹظ…ط© ظˆط§ط®طھط§ط± ط§ظ„ظ…ط§ظƒظٹظ†ط© ط§ظ„طµط­.',
      'Use "Assign machine" from the menu and pick the right machine.'),
  _NoteKind(
      'readfail',
      2,
      Icons.error_outline,
      'ظپط´ظ„ ظپظٹ ظ‚ط±ط§ط،ط© طµظپط­ط§طھ',
      'Pages failed to read',
      'ط§ظ„طھط·ط¨ظٹظ‚ ظ…ط§ ظ‚ط¯ط±ط´ ظٹظ‚ط±ط§ طµظپط­ط© ط£ظˆ ط£ظƒطھط± (ط£ظˆ ط§طھظ‚ظپظ„ ط£ط«ظ†ط§ط، ط§ظ„ط¹ط¯)طŒ ظپط§ظ„ط¹ط¯ط¯ ظ…ظ…ظƒظ† ظٹظƒظˆظ† ظ†ط§ظ‚طµ.',
      'The app could not read one or more pages (or it closed during counting), so the count may be incomplete.',
      'ط§ط¶ط؛ط· "ط¥ط¹ط§ط¯ط© ط§ظ„ط¹ط¯" ظˆط¨ط¹ط¯ظٹظ† "ط¨ط¯ط، ط§ظ„ط¹ط¯". ظ„ظˆ ط§طھظƒط±ط± ط±ط§ط¬ط¹ ط§ظ„ظ…ظ„ظپ ط¨ظ†ظپط³ظƒ.',
      'Tap "Recount" then "Start counting". If it repeats, check the file yourself.'),
  _NoteKind(
      'emptypages',
      1,
      Icons.insert_drive_file_outlined,
      'طµظپط­ط§طھ ظ…ظ† ط؛ظٹط± ط¥ظٹطµط§ظ„ط§طھ',
      'Pages without receipts',
      'ظپظٹظ‡ طµظپط­ط§طھ ظپظٹ ط§ظ„ظ…ظ„ظپ ط§ظ„طھط·ط¨ظٹظ‚ ظ…ط§ ظ„ظ‚ط§ط´ ظپظٹظ‡ط§ ط£ظٹ ط¥ظٹطµط§ظ„ (ظ…ظ…ظƒظ† طھظƒظˆظ† طµظپط­ط© ظپط§ط¶ظٹط© ط£ظˆ طµظˆط±ط© ظ…ط´ ظˆط§ط¶ط­ط©).',
      'Some pages had no receipts detected (they may be blank or unclear).',
      'ط§ظپطھط­ ط§ظ„طµظپط­ط§طھ ط§ظ„ظ…ط°ظƒظˆط±ط© ظˆطھط£ظƒط¯ ط¥ظ†ظ‡ط§ ظپط¹ظ„ط§ظ‹ ظپط§ط¶ظٹط©.',
      'Open the listed pages and make sure they are really empty.'),
  _NoteKind(
      'methoddiff',
      1,
      Icons.difference_outlined,
      'طھط¶ط§ط±ط¨ ط¨ظٹظ† ط·ط±ظٹظ‚طھظٹظ† ظ„ظ„ط¹ط¯',
      'Two counting methods disagree',
      'ط§ظ„طھط·ط¨ظٹظ‚ ط¨ظٹط¹ط¯ ط¨ط·ط±ظٹظ‚طھظٹظ† ظ„ظ„طھط£ظƒط¯طŒ ظˆظپظٹ طµظپط­ط§طھ ظ…ط¹ظٹظ†ط© ط§ظ„ظ†طھظٹط¬ط© ط§ط®طھظ„ظپطھ.',
      'The app counts with two methods to double-check, and on some pages the results differ.',
      'ط±ط§ط¬ط¹ ط§ظ„طµظپط­ط§طھ ط§ظ„ظ…ط°ظƒظˆط±ط© ظ„ظˆ ط§ظ„ط¹ط¯ط¯ ظ…ظ‡ظ… ظ„ظٹظƒ.',
      'Check the listed pages if the exact count matters.'),
  _NoteKind(
      'nonpay',
      0,
      Icons.info_outline,
      'ظ…ط³طھظ†ط¯ ط¥ظ‚ظپط§ظ„ ط£ظˆ ظ…ظ„ط®طµ',
      'Closing or summary document',
      'ط§ظ„ظ…ظ„ظپ ظپظٹظ‡ ط¥ظٹطµط§ظ„ ط¥ظ‚ظپط§ظ„ ط£ظˆ ظ…ظ„ط®طµ ط£ظˆ ط¥ط¹ط§ط¯ط© ط·ط¨ط§ط¹ط©. ط§ظ„طھط·ط¨ظٹظ‚ طھط¬ط§ظ‡ظ„ظ‡ ظˆظ…ط´ ظ…ط­ط³ظˆط¨ ظپظٹ ط§ظ„ط¹ط¯ط¯.',
      'The file contains a closing, summary or reprint receipt. The app ignored it and it is not counted.',
      'ظ…ظپظٹط´ ط­ط§ط¬ط© ظ…ط·ظ„ظˆط¨ط©طŒ ط¯ظ‡ ط·ط¨ظٹط¹ظٹ.',
      'Nothing to do, this is normal.'),
  _NoteKind(
      'nofile',
      2,
      Icons.folder_off_outlined,
      'ط§ظ„ظ…ظ„ظپ ط§ظ„ط£طµظ„ظٹ ظ…ط´ ظ…ظˆط¬ظˆط¯',
      'Original file not found',
      'ط§ظ„ظ…ظ„ظپ ط§طھظ…ط³ط­ ط£ظˆ ط§طھظ†ظ‚ظ„ ظ…ظ† ظ…ظƒط§ظ†ظ‡طŒ ظپط§ظ„طھط·ط¨ظٹظ‚ ظ…ط´ ظ‚ط§ط¯ط± ظٹط¹ط¯ظ‘ظ‡.',
      'The file was deleted or moved, so the app cannot count it.',
      'ط±ط¬ظ‘ط¹ ط§ظ„ظ…ظ„ظپ ظ…ظƒط§ظ†ظ‡ ط£ظˆ ط§ط¹ظ…ظ„ ظپط­طµ ظ…ظ† ط¬ط¯ظٹط¯.',
      'Put the file back or scan again.'),
  _NoteKind(
      'datefail',
      1,
      Icons.calendar_month_outlined,
      'ط§ظ„طھط§ط±ظٹط® ظ…ط´ ظˆط§ط¶ط­',
      'Date not clear',
      'ط§ظ„طھط·ط¨ظٹظ‚ ظ…ط§ ظ‚ط¯ط±ط´ ظٹظ‚ط±ط§ طھط§ط±ظٹط® ط§ظ„ظٹظˆظ… ظ…ظ† ط§ظ„ظ…ظ„ظپ.',
      'The app could not read the date from the file.',
      'ط§ظپطھط­ ط§ظ„ظ…ظ„ظپ ظˆطھط£ظƒط¯ ظ…ظ† ط§ظ„طھط§ط±ظٹط® ط¨ظ†ظپط³ظƒ.',
      'Open the file and check the date yourself.'),
  _NoteKind(
      'namefail',
      2,
      Icons.drive_file_rename_outline,
      'ظپط´ظ„طھ ط§ظ„طھط³ظ…ظٹط©',
      'Naming failed',
      'ط§ظ„طھط·ط¨ظٹظ‚ ظ…ط§ ظ‚ط¯ط±ط´ ظٹطھط¹ط±ظپ ط¹ظ„ظ‰ ط§ظ„ظ…ط§ظƒظٹظ†ط© ظپظ…ط³ظ…ظ‘ط§ط´ ط§ظ„ظ…ظ„ظپ.',
      'The app could not identify the machine so it did not rename the file.',
      'ط§ط³طھط®ط¯ظ… "طھط­ط¯ظٹط¯ ط§ظ„ظ…ط§ظƒظٹظ†ط©" ظ…ظ† ط§ظ„ظ‚ط§ظٹظ…ط©.',
      'Use "Assign machine" from the menu.'),
  _NoteKind(
      'idweak',
      2,
      Icons.help_outline,
      'ط§ظ„ظ…ط§ظƒظٹظ†ط© ظ…ط´ ظ…ط¤ظƒط¯ط©',
      'Machine not confirmed',
      'ط±ظ‚ظ… ط§ظ„ظ…ط§ظƒظٹظ†ط© (Terminal ID) ظ…ط§ ط§طھظ‚ط±ط§ط´ ط¨ظˆط¶ظˆط­طŒ ظپط§ظ„طھط·ط¨ظٹظ‚ ط®ظ…ظ‘ظ† ط§ظ„ظ…ط§ظƒظٹظ†ط© ظ…ظ† ط¨ظٹط§ظ†ط§طھ طھط§ظ†ظٹط© ط£ظˆ ظ„ظ‚ط§ظ‡ط§ ظپظٹ ط£ظƒطھط± ظ…ظ† ظ…ظ„ظپ.',
      'The Terminal ID was not read clearly, so the app guessed the machine from other data or found it in more than one file.',
      'ط§طھط£ظƒط¯ ظ…ظ† ط§ظ„ظ…ط§ظƒظٹظ†ط©طŒ ظˆظ„ظˆ ط؛ظ„ط· ط§ط³طھط®ط¯ظ… "طھط­ط¯ظٹط¯ ط§ظ„ظ…ط§ظƒظٹظ†ط©".',
      'Verify the machine; if wrong, use "Assign machine".'),
  _NoteKind(
      'other',
      1,
      Icons.report_problem_outlined,
      'ظ…ظ„ط§ط­ط¸ط© طھط§ظ†ظٹط©',
      'Other note',
      'ظ…ظ„ط§ط­ط¸ط© ظ…ظ† ط§ظ„طھط·ط¨ظٹظ‚ ظ…ط´ ظ…طھطµظ†ظپط©.',
      'An unclassified note from the app.',
      'ط§ظ‚ط±ط§ظ‡ط§ ظˆظ‚ط±ط± ظ„ظˆ ظ…ط­طھط§ط¬ط© ظ…ط±ط§ط¬ط¹ط©.',
      'Read it and decide whether it needs review.'),
];

final List<List<Object>> _noteRules = [
  [RegExp(r'^ط§ظ„ط¹ط¯ط¯ \d+ ط¨ط³ ط§ظ„طھط³ظ„ط³ظ„'), 'seq'],
  [RegExp(r'ظ…ط´ طھط¨ط¹ ظ†ظپط³ ط§ظ„ظٹظˆظ…'), 'otherday'],
  [RegExp(r'ط§ظ„ط¹ط¯ ظ„ظ‚ظ‰ ظ…ط§ظƒظٹظ†ط©'), 'machconf'],
  [RegExp(r'طµظپط­ط© ظپط´ظ„طھ ظ‚ط±ط§ط،طھظ‡ط§|ظپط´ظ„ ط§ظ„ط¹ط¯|ط§ظ„ط¹ط¯ ظˆظ‚ظپ'), 'readfail'],
  [RegExp(r'طµظپط­ط§طھ ظ…ط§ ظ„ظ‚ظٹطھط´ ظپظٹظ‡ط§ ط¥ظٹطµط§ظ„ط§طھ'), 'emptypages'],
  [RegExp(r'ط§ظ„ط·ط±ظٹظ‚طھظٹظ† ط§ط®طھظ„ظپظˆط§'), 'methoddiff'],
  [RegExp(r'ط¥ظ‚ظپط§ظ„/ظ…ظ„ط®طµ'), 'nonpay'],
  [RegExp(r'ط§ظ„ظ…ظ„ظپ ط§ظ„ط£طµظ„ظٹ ظ…ط´ ظ…ظˆط¬ظˆط¯'), 'nofile'],
  [RegExp(r'ظ…ط´ ظ‚ط§ط¯ط± ط£ط¹ط±ظپ ط§ظ„طھط§ط±ظٹط®'), 'datefail'],
  [RegExp(r'ظپط´ظ„ ط§ظ„طھط³ظ…ظٹط©|ظپط´ظ„ ط§ظ„طھط¹ط±ظٹظپ'), 'namefail'],
  [RegExp(r'Terminal ID|ظ…ط´ ظˆط§ط¶ط­|ظ…ط´ ظ‚ط§ط¯ط± ط£ط¹ط±ظپ ط§ظ„ظ…ط§ظƒظٹظ†ط©|ظ†ظپط³ ظ…ط§ظƒظٹظ†ط©|ط§طھط¹ط±ظپطھ ظƒظ…ط§ظ†|ط§ظ„ظˆط­ط¯ط© ط§طھط­ط¯ط¯طھ'), 'idweak'],
];

_NoteKind _noteKindOf(String n) {
  for (final r in _noteRules) {
    if ((r[0] as RegExp).hasMatch(n)) {
      return _kNoteKinds.firstWhere((k) => k.key == r[1]);
    }
  }
  return _kNoteKinds.last;
}

class _MStats {
  int? pages, full, partial, seqMin, seqMax, seqLen;
}

_MStats _statsOf(MResult r) {
  final s = _MStats();
  final d = r.detail;
  int? g(RegExp re) {
    final m = re.firstMatch(d);
    return m == null ? null : int.tryParse(m.group(1)!);
  }

  s.pages = g(RegExp(r'(\d+)\s*طµظپط­ط©'));
  s.full = g(RegExp(r'RRN\s*(?:ظ…ظ‚ط±ظˆط،|ظƒط§ظ…ظ„)\s*(\d+)'));
  s.partial = g(RegExp(r'ظ†ط§ظ‚طµ ط§ظ„ظ‚ط±ط§ط،ط©\s*(\d+)'));
  final m = RegExp(r'ط§ظ„طھط³ظ„ط³ظ„\s*(\d+)-(\d+)\s*\((\d+)\)').firstMatch(d);
  if (m != null) {
    s.seqMin = int.tryParse(m.group(1)!);
    s.seqMax = int.tryParse(m.group(2)!);
    s.seqLen = int.tryParse(m.group(3)!);
  }
  return s;
}

class _St {
  final String ar, en;
  final MaterialColor color;
  const _St(this.ar, this.en, this.color);
}

_St _statusOf(MResult r) {
  if (r.excluded) return const _St('ظ…ط³طھط¨ط¹ط¯', 'Excluded', Colors.grey);
  if (r.pending) return const _St('ظ„ط³ظ‡ ظ…ط§ ط§طھط¹ط¯ظ‘ط´', 'Not counted yet', Colors.blueGrey);
  if (r.verdict > 0) return const _St('ط§طھط±ط§ط¬ط¹: طµط­', 'Reviewed: OK', Colors.green);
  if (r.ok) return const _St('ط³ظ„ظٹظ…', 'OK', Colors.green);
  return const _St('ظ…ط­طھط§ط¬ ظ…ط±ط§ط¬ط¹ط©', 'Needs review', Colors.orange);
}

// ط§ظ„ط¨ط§طھط´ = ط§ظ„ط¹ط¯ط¯ ط§ظ„ظ„ظٹ ظ…ظƒطھظˆط¨ ظپظٹ ط®ط§ظ†ط© "ط§ظ„طµط­" (ظ…ظ† ط¥ظٹطµط§ظ„ ط§ظ„ط¨ط§طھط´)طŒ ظˆط§ظ„ظپط¹ظ„ظٹ = ط§ظ„ط¹ط¯ط¯ ط§ظ„ظ„ظٹ ط§ظ„طھط·ط¨ظٹظ‚ ط¹ط¯ظ‘ظ‡
int? _batchManual(MResult r) => int.tryParse(r.ctl.text.trim());
int? _batchAutoOf(MResult r) => r.orderNum == null ? null : closingTxn.value[r.orderNum!];
int? _batchOf(MResult r) => _batchManual(r) ?? _batchAutoOf(r);
bool _batchIsAuto(MResult r) => _batchManual(r) == null && _batchAutoOf(r) != null;

// ظ…ط·ط§ط¨ظ‚ط© ط§ظ„ط¨ط§طھط´: ظ„ظˆ ط§ظ†طھ ط±ط§ط¬ط¹طھ ط§ظ„ظ…ظ„ظپ ظٹط¯ظˆظٹ ظˆط¹ظ„ظ‘ظ…طھظ‡ طµط­ âœ“ ظٹطھط­ط³ط¨ ظ…ط·ط§ط¨ظ‚ (ط£ط®ط¶ط±) ط­طھظ‰ ظ„ظˆ ط§ظ„ط¹ط¯ط¯ ط§ظ„ظ…ظ‚ط±ظˆط، ظ…ط®طھظ„ظپ ط¹ظ† ط±ظ‚ظ… ط§ظ„ط¨ط§طھط´.
// null = ظ…ظپظٹط´ ط­ط§ط¬ط© ظ†ظ‚ط§ط±ظ†ظ‡ط§. ط§ظ„ط±ظ‚ظ… ط§ظ„ظ„ظٹ ط¨طھظƒطھط¨ظ‡ ط¨ط¥ظٹط¯ظƒ ظپظٹ ط®ط§ظ†ط© "ط§ظ„طµط­" ط¨ظٹط؛ظ„ط¨ ط§ظ„ط±ظ‚ظ… ط§ظ„طھظ„ظ‚ط§ط¦ظٹ ظ…ظ† ط§ظ„طھظ‚ظپظٹظ„.
bool? _batchMatch(MResult r) {
  if (r.excluded || r.pending) return null;
  if (r.verdict > 0) return true;
  final d = _batchDiff(r);
  return d == null ? null : d == 0;
}

int? _batchDiff(MResult r) {
  final b = _batchOf(r);
  if (b == null || r.pending || r.excluded) return null;
  return b - r.count;
}

String _diffText(BuildContext c, int d) => d == 0
    ? _l(c, 'ظ…ط·ط§ط¨ظ‚ âœ“', 'Match âœ“')
    : (d > 0 ? _l(c, 'ظ†ط§ظ‚طµ $d', 'Missing $d') : _l(c, 'ط²ظٹط§ط¯ط© ${-d}', 'Extra ${-d}'));

MaterialColor _diffColor(int d) => d == 0 ? Colors.green : (d > 0 ? Colors.red : Colors.orange);

String _mSummary(BuildContext c, MResult r) {
  if (r.excluded) return _l(c, 'ظ…ط³طھط¨ط¹ط¯ (ط¹ظ„ط§ظ…ط© ط؛ظ„ط·)', 'Excluded (marked wrong)');
  if (r.pending) return _l(c, 'ظ„ط³ظ‡ ظ…ط§ ط§طھط¹ط¯ظ‘ط´ (ط§ط¶ط؛ط· ط¨ط¯ط، ط§ظ„ط¹ط¯)', 'Not counted yet (tap Start counting)');
  final b = StringBuffer('${_l(c, 'ط¹ط¯ط¯ ط§ظ„ظ…ط¹ط§ظ…ظ„ط§طھ', 'Transactions')}: ${r.count}');
  final bt = _batchOf(r);
  final d = _batchDiff(r);
  if (bt != null && d != null) {
    final tag = _batchIsAuto(r)
        ? _l(c, ' (ظ…ظ† ط§ظ„طھظ‚ظپظٹظ„)', ' (from closing)')
        : (_batchManual(r) != null ? _l(c, ' (ظٹط¯ظˆظٹ)', ' (manual)') : '');
    final res = (r.verdict > 0 && d != 0)
        ? _l(c, 'ط±ط§ط¬ط¹طھظ‡ ظٹط¯ظˆظٹ: طµط­ âœ“', 'Manually reviewed: OK âœ“')
        : _diffText(c, d);
    b.write('\n${_l(c, 'ط§ظ„ط¨ط§طھط´', 'Batch')}: $bt$tag â€¢ ${_l(c, 'ط§ظ„ظپط¹ظ„ظٹ', 'Actual')}: ${r.count} â€¢ $res');
  }
  return b.toString();
}

Widget _statusChip(BuildContext c, _St st) => Container(
      padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 3),
      decoration: BoxDecoration(
        color: _tintBg(c, st.color),
        borderRadius: BorderRadius.circular(20),
        border: Border.all(color: _tintBorder(c, st.color)),
      ),
      child: w.Text(_l(c, st.ar, st.en),
          style: TextStyle(fontSize: 12, fontWeight: FontWeight.bold, color: _tintFg(c, st.color))),
    );

class ReceiptReportPage extends StatefulWidget {
  final MResult r;
  final bool canEdit;
  final VoidCallback onChanged;
  final VoidCallback onOpenPdf;
  final VoidCallback onRecount;
  final void Function(int) onVerdict;
  const ReceiptReportPage({
    super.key,
    required this.r,
    required this.canEdit,
    required this.onChanged,
    required this.onOpenPdf,
    required this.onRecount,
    required this.onVerdict,
  });

  @override
  State<ReceiptReportPage> createState() => _ReceiptReportPageState();
}

class _ReceiptReportPageState extends State<ReceiptReportPage> {
  @override
  Widget build(BuildContext context) {
    AppSettings.of(context);
    final r = widget.r;
    final cs = Theme.of(context).colorScheme;
    final st = _statusOf(r);
    final stats = _statsOf(r);
    final bd = _batchDiff(r);
    final bt = _batchOf(r);
    final can = widget.canEdit;
    final groups = <String, List<String>>{};
    for (final n in r.notes) {
      groups.putIfAbsent(_noteKindOf(n).key, () => <String>[]).add(n);
    }
    final kinds = groups.keys.map((k) => _kNoteKinds.firstWhere((e) => e.key == k)).toList()
      ..sort((a, b) => b.sev.compareTo(a.sev));

    Widget section(String t, List<Widget> kids) => Card(
          margin: const EdgeInsets.only(bottom: 8),
          child: Padding(
            padding: const EdgeInsets.all(12),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.stretch,
              children: [
                w.Text(t, style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 15)),
                const SizedBox(height: 8),
                ...kids,
              ],
            ),
          ),
        );
    Widget kv(String k, String v, {Color? color}) => Padding(
          padding: const EdgeInsets.symmetric(vertical: 3),
          child: Row(
            children: [
              Expanded(child: w.Text(k, style: TextStyle(fontSize: 13, color: cs.onSurfaceVariant))),
              w.Text(v, style: TextStyle(fontSize: 14, fontWeight: FontWeight.bold, color: color)),
            ],
          ),
        );
    String pad6(int? n) => (n ?? 0).toString().padLeft(6, '0');

    final numbers = <Widget>[
      kv(_l(context, 'ط¹ط¯ط¯ ط§ظ„ظ…ط¹ط§ظ…ظ„ط§طھ ط§ظ„ظ…ط¹ط¯ظˆط¯ط©', 'Counted transactions'), r.pending ? 'â€”' : '${r.count}'),
      if (stats.pages != null) kv(_l(context, 'ط¹ط¯ط¯ ط§ظ„طµظپط­ط§طھ', 'Pages'), '${stats.pages}'),
      if (stats.full != null) kv(_l(context, 'ط¥ظٹطµط§ظ„ط§طھ ظ…ظ‚ط±ظˆط،ط© ظƒط§ظ…ظ„ط©', 'Fully read receipts'), '${stats.full}'),
      if (stats.partial != null)
        kv(_l(context, 'ط¥ظٹطµط§ظ„ط§طھ ظ†ط§ظ‚طµط© ط§ظ„ظ‚ط±ط§ط،ط©', 'Partially read receipts'), '${stats.partial}',
            color: stats.partial! > 0 ? Colors.orange : null),
      if (stats.seqMin != null)
        kv(_l(context, 'طھط³ظ„ط³ظ„ ط§ظ„ط¥ظٹطµط§ظ„ط§طھ', 'Receipt sequence'), '${pad6(stats.seqMin)} â†’ ${pad6(stats.seqMax)}'),
      if (stats.seqLen != null)
        kv(_l(context, 'ط§ظ„ط¹ط¯ط¯ ط§ظ„ظ…طھظˆظ‚ط¹ ظ…ظ† ط§ظ„طھط³ظ„ط³ظ„', 'Expected from sequence'), '${stats.seqLen}'),
      if (stats.seqLen != null && !r.pending)
        kv(
            _l(context, 'ط§ظ„ظپط±ظ‚ ط¹ظ† ط§ظ„طھط³ظ„ط³ظ„', 'Difference from sequence'),
            r.count == stats.seqLen ? _l(context, 'ظ…ط·ط§ط¨ظ‚ âœ“', 'Match âœ“') : '${r.count - stats.seqLen!}',
            color: r.count == stats.seqLen ? Colors.green : Colors.orange),
    ];

    return Scaffold(
      appBar: AppBar(title: w.Text(_l(context, 'طھظ‚ط±ظٹط± ط§ظ„ط¥ظٹطµط§ظ„', 'Receipt report'))),
      body: ListView(
        padding: const EdgeInsets.all(8),
        children: [
          Card(
            margin: const EdgeInsets.only(bottom: 8),
            child: Padding(
              padding: const EdgeInsets.all(12),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Row(
                    children: [
                      Expanded(
                        child: Text(r.newName ?? 'طں ${r.pdfName}',
                            style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 16)),
                      ),
                      const SizedBox(width: 8),
                      _statusChip(context, st),
                    ],
                  ),
                  const SizedBox(height: 4),
                  w.Text('${_l(context, 'ط§ظ„ظ…ظ„ظپ ط§ظ„ط£طµظ„ظٹ', 'Original file')}: ${r.pdfName}',
                      style: TextStyle(fontSize: 12, color: cs.onSurfaceVariant)),
                ],
              ),
            ),
          ),
          section(_l(context, 'ط§ظ„ط£ط±ظ‚ط§ظ…', 'Numbers'), [
            ...numbers,
            if (r.detail.isNotEmpty)
              Padding(
                padding: const EdgeInsets.only(top: 6),
                child: Text(r.detail, style: TextStyle(fontSize: 11, color: cs.onSurfaceVariant)),
              ),
          ]),
          section(_l(context, 'ط±ط¨ط· ط§ظ„ط¨ط§طھط´', 'Batch link'), [
            TextField(
              controller: r.ctl,
              keyboardType: TextInputType.number,
              decoration: InputDecoration(
                isDense: true,
                labelText: _l(context, 'ط¹ط¯ط¯ ظ…ط¹ط§ظ…ظ„ط§طھ ط§ظ„ط¨ط§طھط´ (ط¨ظٹطھظ‚ط±ظٹ ظ…ظ† ط§ظ„طھظ‚ظپظٹظ„طŒ ط§ظƒطھط¨ ط±ظ‚ظ… ظ„ظˆ ط¹ط§ظٹط² طھط؛ظٹظ‘ط±ظ‡)',
                    'Batch transaction count (read from closing; type a number to override)'),
                hintText: _batchAutoOf(r)?.toString(),
              ),
              onChanged: (_) {
                widget.onChanged();
                setState(() {});
              },
            ),
            const SizedBox(height: 8),
            if (bd == null || bt == null)
              w.Text(
                  _l(context, 'ظ…ظپظٹط´ ط¹ط¯ط¯ ط¨ط§طھط´ ط§طھظ‚ط±ط§ ظ…ظ† ط§ظ„طھظ‚ظپظٹظ„ ظ„ظ…ط§ظƒظٹظ†ط© ط¯ظٹ. ط§ظƒطھط¨ظ‡ ظ‡ظ†ط§ ط¹ظ„ط´ط§ظ† ظ†ظ‚ط§ط±ظ†ظ‡ ط¨ط§ظ„ط¹ط¯ط¯ ط§ظ„ظپط¹ظ„ظٹ.',
                      'No batch count was read from closing for this machine. Enter it here to compare with the actual count.'),
                  style: TextStyle(fontSize: 12, color: cs.onSurfaceVariant))
            else
              Container(
                padding: const EdgeInsets.all(10),
                decoration: BoxDecoration(
                  color: _tintBg(context, r.verdict > 0 ? Colors.green : _diffColor(bd), soft: true),
                  border: Border.all(color: _tintBorder(context, r.verdict > 0 ? Colors.green : _diffColor(bd))),
                  borderRadius: BorderRadius.circular(10),
                ),
                child: Column(
                  children: [
                    kv(_l(context, 'ط§ظ„ط¨ط§طھط´', 'Batch'),
                        '$bt${_batchIsAuto(r) ? _l(context, ' (ظ…ظ† ط§ظ„طھظ‚ظپظٹظ„)', ' (from closing)') : ''}'),
                    kv(_l(context, 'ط§ظ„ظپط¹ظ„ظٹ', 'Actual'), '${r.count}'),
                    kv(
                        _l(context, 'ط§ظ„ظپط±ظ‚', 'Difference'),
                        '${_diffText(context, bd)}${r.verdict > 0 && bd != 0 ? _l(context, ' â€¢ ط±ط§ط¬ط¹طھظ‡ ظٹط¯ظˆظٹ: طµط­ âœ“', ' â€¢ manually reviewed: OK âœ“') : ''}',
                        color: _tintFg(context, r.verdict > 0 ? Colors.green : _diffColor(bd))),
                  ],
                ),
              ),
          ]),
          section(_l(context, 'ط§ظ„ظ…ظ„ط§ط­ط¸ط§طھ', 'Notes'), [
            if (kinds.isEmpty)
              Container(
                padding: const EdgeInsets.all(10),
                decoration: BoxDecoration(
                  color: _tintBg(context, Colors.green, soft: true),
                  border: Border.all(color: _tintBorder(context, Colors.green)),
                  borderRadius: BorderRadius.circular(10),
                ),
                child: w.Text(_l(context, 'ظ…ظپظٹط´ ظ…ظ„ط§ط­ط¸ط§طھ âœ“', 'No notes âœ“'),
                    style: TextStyle(fontWeight: FontWeight.bold, color: _tintFg(context, Colors.green))),
              ),
            for (final k in kinds)
              Builder(builder: (context) {
                final MaterialColor col = k.sev == 2 ? Colors.red : (k.sev == 1 ? Colors.orange : Colors.blue);
                final sevLabel = k.sev == 2
                    ? _l(context, 'ظ„ط§ط²ظ… طھط±ط§ط¬ط¹', 'Must review')
                    : (k.sev == 1 ? _l(context, 'ط§ظ†طھط¨ظ‡', 'Heads up') : _l(context, 'ظ…ط¹ظ„ظˆظ…ط©', 'Info'));
                return Container(
                  margin: const EdgeInsets.only(bottom: 8),
                  padding: const EdgeInsets.all(10),
                  decoration: BoxDecoration(
                    color: _tintBg(context, col, soft: true),
                    border: Border.all(color: _tintBorder(context, col)),
                    borderRadius: BorderRadius.circular(10),
                  ),
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Row(
                        children: [
                          Icon(k.icon, size: 20, color: _tintFg(context, col)),
                          const SizedBox(width: 8),
                          Expanded(
                            child: w.Text(_l(context, k.titleAr, k.titleEn),
                                style: TextStyle(fontWeight: FontWeight.bold, color: _tintFg(context, col))),
                          ),
                          w.Text(sevLabel, style: TextStyle(fontSize: 11, color: _tintFg(context, col))),
                        ],
                      ),
                      const SizedBox(height: 6),
                      w.Text(_l(context, k.whyAr, k.whyEn), style: const TextStyle(fontSize: 13)),
                      const SizedBox(height: 6),
                      for (final n in groups[k.key]!)
                        Padding(
                          padding: const EdgeInsets.only(bottom: 2),
                          child: Text('â€¢ $n', style: const TextStyle(fontSize: 12)),
                        ),
                      const SizedBox(height: 6),
                      w.Text('${_l(context, 'طھط¹ظ…ظ„ ط¥ظٹظ‡طں', 'What to do')}: ${_l(context, k.fixAr, k.fixEn)}',
                          style: const TextStyle(fontSize: 13, fontWeight: FontWeight.bold)),
                    ],
                  ),
                );
              }),
          ]),
          Wrap(
            spacing: 6,
            runSpacing: 6,
            children: [
              _mini(Icons.picture_as_pdf, _l(context, 'ظپطھط­ ط§ظ„ظ…ظ„ظپ', 'Open file'), widget.onOpenPdf, filled: true),
              _mini(Icons.check, _l(context, 'طµط­ âœ“', 'OK âœ“'), !can
                  ? null
                  : () {
                      widget.onVerdict(1);
                      setState(() {});
                    }),
              _mini(Icons.close, _l(context, 'ط؛ظ„ط· âœ—', 'Wrong âœ—'), !can
                  ? null
                  : () {
                      widget.onVerdict(-1);
                      setState(() {});
                    }),
              if (r.verdict != 0)
                _mini(Icons.undo, _l(context, 'ط¥ظ„ط؛ط§ط، ط§ظ„ط¹ظ„ط§ظ…ط©', 'Clear mark'), !can
                    ? null
                    : () {
                        widget.onVerdict(0);
                        setState(() {});
                      }),
              _mini(Icons.refresh, _l(context, 'ط¥ط¹ط§ط¯ط© ط§ظ„ط¹ط¯', 'Recount'), !can
                  ? null
                  : () {
                      widget.onRecount();
                      setState(() {});
                      ScaffoldMessenger.of(context).showSnackBar(SnackBar(
                          content: w.Text(_l(context, 'ط§طھط¹ظ„ظ‘ظ… ظ„ظ„ط¹ط¯. ط§ط±ط¬ط¹ ظˆط§ط¶ط؛ط· "ط¨ط¯ط، ط§ظ„ط¹ط¯".',
                              'Marked for counting. Go back and tap "Start counting".'))));
                    }),
            ],
          ),
          const SizedBox(height: 16),
        ],
      ),
    );
  }
}

class MachinesReportPage extends StatefulWidget {
  final List<MResult> results;
  final List<String> missing;
  final Future<void> Function(MResult) openReport;
  const MachinesReportPage({
    super.key,
    required this.results,
    required this.missing,
    required this.openReport,
  });

  @override
  State<MachinesReportPage> createState() => _MachinesReportPageState();
}

class _MachinesReportPageState extends State<MachinesReportPage> {
  int filter = 0; // 0 ط§ظ„ظƒظ„طŒ 1 ظ…ط±ط§ط¬ط¹ط©طŒ 2 ظپط±ظˆظ‚ ط§ظ„ط¨ط§طھط´

  @override
  Widget build(BuildContext context) {
    AppSettings.of(context);
    final cs = Theme.of(context).colorScheme;
    final all = widget.results;
    final counted = all.where((r) => !r.excluded && !r.pending).toList();
    final totalTx = counted.fold<int>(0, (a, r) => a + r.count);
    final reviewN = all.where((r) => !r.ok).length;
    final excl = all.where((r) => r.excluded).length;
    final pend = all.where((r) => r.pending).length;
    final linked = counted.where((r) => _batchOf(r) != null).toList();
    final matched = linked.where((r) => _batchMatch(r) == true).length;
    final diffs = linked.length - matched;
    final batchSum = linked.fold<int>(0, (a, r) => a + _batchOf(r)!);
    final actualSum = linked.fold<int>(0, (a, r) => a + r.count);
    final shown = all
        .where((r) => filter == 0 || (filter == 1 && !r.ok) || (filter == 2 && _batchMatch(r) == false))
        .toList();

    Widget stat(String label, String value, MaterialColor c) => Container(
          width: 108,
          padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 10),
          decoration: BoxDecoration(
            color: _tintBg(context, c, soft: true),
            border: Border.all(color: _tintBorder(context, c)),
            borderRadius: BorderRadius.circular(10),
          ),
          child: Column(
            children: [
              w.Text(value, style: TextStyle(fontSize: 20, fontWeight: FontWeight.bold, color: _tintFg(context, c))),
              const SizedBox(height: 2),
              w.Text(label,
                  textAlign: TextAlign.center, style: TextStyle(fontSize: 11, color: _tintFg(context, c))),
            ],
          ),
        );
    Widget kv(String k, String v, {Color? color}) => Padding(
          padding: const EdgeInsets.symmetric(vertical: 3),
          child: Row(
            children: [
              Expanded(child: w.Text(k, style: TextStyle(fontSize: 13, color: cs.onSurfaceVariant))),
              w.Text(v, style: TextStyle(fontSize: 14, fontWeight: FontWeight.bold, color: color)),
            ],
          ),
        );

    return Scaffold(
      appBar: AppBar(title: w.Text(_l(context, 'طھظ‚ط±ظٹط± ط¥ظٹطµط§ظ„ط§طھ ط§ظ„ط¯ظپط¹', 'Payment receipts report'))),
      body: ListView(
        padding: const EdgeInsets.all(8),
        children: [
          Wrap(
            spacing: 8,
            runSpacing: 8,
            children: [
              stat(_l(context, 'ط¥ط¬ظ…ط§ظ„ظٹ ط§ظ„ظ…ط¹ط§ظ…ظ„ط§طھ', 'Total transactions'), '$totalTx', Colors.blue),
              stat(_l(context, 'ظ…ظ„ظپط§طھ', 'Files'), '${all.length}', Colors.blueGrey),
              stat(_l(context, 'ط§طھط¹ط¯ظ‘طھ', 'Counted'), '${counted.length}', Colors.green),
              stat(_l(context, 'ظ…ط­طھط§ط¬ ظ…ط±ط§ط¬ط¹ط©', 'Need review'), '$reviewN', reviewN > 0 ? Colors.orange : Colors.green),
              stat(_l(context, 'ظ…ط³طھط¨ط¹ط¯', 'Excluded'), '$excl', Colors.grey),
              stat(_l(context, 'ظ„ط³ظ‡ ظ…ط§ ط§طھط¹ط¯ظ‘طھط´', 'Not counted'), '$pend', pend > 0 ? Colors.amber : Colors.green),
            ],
          ),
          const SizedBox(height: 8),
          Card(
            margin: EdgeInsets.zero,
            child: Padding(
              padding: const EdgeInsets.all(12),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.stretch,
                children: [
                  w.Text(_l(context, 'ط±ط¨ط· ط§ظ„ط¨ط§طھط´', 'Batch link'),
                      style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 15)),
                  const SizedBox(height: 6),
                  if (linked.isEmpty)
                    w.Text(
                        _l(context, 'ط¹ط¯ط¯ ط§ظ„ط¨ط§طھط´ ط¨ظٹطھظ‚ط±ظٹ طھظ„ظ‚ط§ط¦ظٹ ظ…ظ† ط¥ظٹطµط§ظ„ط§طھ طھط§ط¨ ط§ظ„طھظ‚ظپظٹظ„. ظ„ظˆ ظ…ط§ ط¸ظ‡ط±ط´ ط§ظƒطھط¨ظ‡ ظپظٹ ط®ط§ظ†ط© "ط§ظ„طµط­" ط£ظˆ ظپظٹ طھظ‚ط±ظٹط± ط§ظ„ط¥ظٹطµط§ظ„.',
                            'The batch count is read automatically from the closing tab receipts. If it is missing, type it in the "Correct" field or in the receipt report.'),
                        style: TextStyle(fontSize: 12, color: cs.onSurfaceVariant))
                  else ...[
                    kv(_l(context, 'ظ…ط§ظƒظٹظ†ط§طھ ظ…ط±ط¨ظˆط·ط©', 'Linked machines'), '${linked.length} / ${counted.length}'),
                    kv(_l(context, 'ظ…ط·ط§ط¨ظ‚ط©', 'Matching'), '$matched', color: Colors.green),
                    kv(_l(context, 'ظپظٹظ‡ط§ ظپط±ظ‚', 'With a difference'), '$diffs',
                        color: diffs > 0 ? Colors.red : Colors.green),
                    kv(_l(context, 'ط¥ط¬ظ…ط§ظ„ظٹ ط§ظ„ط¨ط§طھط´ط§طھ', 'Batch total'), '$batchSum'),
                    kv(_l(context, 'ط¥ط¬ظ…ط§ظ„ظٹ ط§ظ„ظپط¹ظ„ظٹ', 'Actual total'), '$actualSum'),
                    kv(_l(context, 'طµط§ظپظٹ ط§ظ„ظپط±ظ‚', 'Net difference'), _diffText(context, batchSum - actualSum),
                        color: _tintFg(context, _diffColor(batchSum - actualSum))),
                  ],
                ],
              ),
            ),
          ),
          if (widget.missing.isNotEmpty) ...[
            const SizedBox(height: 8),
            Container(
              padding: const EdgeInsets.all(10),
              decoration: BoxDecoration(
                color: _tintBg(context, Colors.amber),
                border: Border.all(color: _tintBorder(context, Colors.amber)),
                borderRadius: BorderRadius.circular(8),
              ),
              child: w.Text(
                  'âڑ ï¸ڈ ${_l(context, 'ظ…ط§ظƒظٹظ†ط§طھ ظ…ط§ ظˆطµظ„ط´ ظ…ظ„ظپظ‡ط§', 'Machines with no file')} (${widget.missing.length}): ${widget.missing.join(' طŒ ')}',
                  style: const TextStyle(fontSize: 13)),
            ),
          ],
          const SizedBox(height: 8),
          Wrap(
            spacing: 6,
            children: [
              for (final o in [
                [0, _l(context, 'ط§ظ„ظƒظ„', 'All')],
                [1, _l(context, 'ظ…ط±ط§ط¬ط¹ط© ($reviewN)', 'Review ($reviewN)')],
                [2, _l(context, 'ظپط±ظˆظ‚ ط§ظ„ط¨ط§طھط´ ($diffs)', 'Batch differences ($diffs)')],
              ])
                ChoiceChip(
                  label: w.Text(o[1] as String, style: const TextStyle(fontSize: 13)),
                  visualDensity: VisualDensity.compact,
                  selected: filter == o[0],
                  onSelected: (_) => setState(() => filter = o[0] as int),
                ),
            ],
          ),
          const SizedBox(height: 6),
          for (final r in shown)
            Card(
              margin: const EdgeInsets.symmetric(vertical: 2),
              child: ListTile(
                contentPadding: const EdgeInsets.symmetric(horizontal: 12),
                onTap: () async {
                  await widget.openReport(r);
                  if (mounted) setState(() {});
                },
                title: Text(r.newName ?? 'طں ${r.pdfName}', style: const TextStyle(fontWeight: FontWeight.bold)),
                subtitle: w.Text(_mSummary(context, r)),
                trailing: _statusChip(context, _statusOf(r)),
              ),
            ),
          if (shown.isEmpty)
            Padding(
              padding: const EdgeInsets.all(16),
              child: Center(child: w.Text(_l(context, 'ظ…ظپظٹط´ ظ…ظ„ظپط§طھ ظپظٹ ط§ظ„ظپظ„طھط± ط¯ظ‡.', 'No files in this filter.'))),
            ),
          const SizedBox(height: 16),
        ],
      ),
    );
  }
}

class MachinesPage extends StatefulWidget {
  const MachinesPage({super.key});

  @override
  State<MachinesPage> createState() => _MachinesPageState();
}

class _MachinesPageState extends State<MachinesPage> with WidgetsBindingObserver {
  String? pdfDir; // null = ظ…ط³طھظ†ط¯ط§طھ ظˆط§طھط³ط§ط¨ (طھظ„ظ‚ط§ط¦ظٹ)
  int days = 0;
  bool skipDone = true;
  bool reviewOnly = false;
  bool namingDone = false; // ط§ظ„طھط³ظ…ظٹط© ط®ظ„طµطھ (ط­طھظ‰ ظ„ظˆ ط§ظ„ط¹ط¯ ظ„ط³ظ‡ ط´ط؛ط§ظ„): ط²ط±ط§ط± ط§ظ„ط¥ط±ط³ط§ظ„ ظٹطھظپط¹ظ„
  bool busy = false;
  bool stopReq = false;
  double progressValue = 0;
  String status = 'ط§ط¶ط؛ط· "ظپط­طµ ظˆط¥ط¹ط§ط¯ط© طھط³ظ…ظٹط©" ظ„ظپط­طµ ط¥ظٹطµط§ظ„ط§طھ ط§ظ„ط¯ظپط¹ ظ…ظ† ظ…ط³طھظ†ط¯ط§طھ ظˆط§طھط³ط§ط¨.';
  final List<MResult> results = [];
  AppLifecycleState _life = AppLifecycleState.resumed;
  String destDir = loadMOut();

  @override
  void initState() {
    super.initState();
    WidgetsBinding.instance.addObserver(this);
    pdfDir = loadPdfDir();
    days = sharedDays.value;
    sharedDays.addListener(_syncDays);
    wipeSignal.addListener(_onWipe);
    closingTxn.addListener(_onClosing);
    results.addAll(loadMState());
    results.sort(_cmpM);
    final pend = results.where((r) => r.pending).length;
    if (results.isNotEmpty) {
      status = pend > 0
          ? 'ظپظٹظ‡ $pend ظ…ظ„ظپ ظ„ط³ظ‡ ظ…ط§ ط§طھط¹ط¯ظ‘ط´. ط§ط¶ط؛ط· \"ط¨ط¯ط، ط§ظ„ط¹ط¯\" ظ„ظ…ط§ طھط­ط¨.'
          : 'ط¢ط®ط± ظ†طھظٹط¬ط©: ${results.length} ظ…ظ„ظپ.';
    }
  }

  @override
  void dispose() {
    sharedDays.removeListener(_syncDays);
    wipeSignal.removeListener(_onWipe);
    closingTxn.removeListener(_onClosing);
    WidgetsBinding.instance.removeObserver(this);
    for (final r in results) {
      r.ctl.dispose();
    }
    super.dispose();
  }

  @override
  void didChangeAppLifecycleState(AppLifecycleState state) {
    _life = state;
    if (busy) _dbg('LIFECYCLE $state');
  }

  void _syncDays() {
    if (!mounted || busy || days == sharedDays.value) return;
    setState(() => days = sharedDays.value);
  }

  void _onWipe() {
    if (!mounted) return;
    for (final r in results) {
      r.ctl.dispose();
    }
    setState(() {
      results.clear();
      progressValue = 0;
      namingDone = false;
      reviewOnly = false;
      status = 'طھظ… ظ…ط³ط­ ظƒظ„ ط­ط§ط¬ط©. ط§ط¶ط؛ط· "ظپط­طµ ظˆط¥ط¹ط§ط¯ط© طھط³ظ…ظٹط©" ظ„ط¨ط¯ط، ظٹظˆظ… ط¬ط¯ظٹط¯.';
    });
  }

  void _onClosing() {
    if (mounted) setState(() {});
  }

  Future<void> openPdf(String? path) async {
    if (path == null || !File(path).existsSync()) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('ط§ظ„ظ…ظ„ظپ ظ…ط´ ظ…ظˆط¬ظˆط¯.')),
      );
      return;
    }
    try {
      final ok = await _chan.invokeMethod<bool>('openFile', {'path': path, 'mime': 'application/pdf'});
      if (ok == true) return;
    } catch (_) {}
    if (!mounted) return;
    ScaffoldMessenger.of(context).showSnackBar(
      const SnackBar(content: Text('ظ…ظپظٹط´ ط¨ط±ظ†ط§ظ…ط¬ ظٹظپطھط­ ط§ظ„ظ€ PDF.')),
    );
  }

  Future<void> pickDir() async {
    final perm = await Permission.manageExternalStorage.request();
    if (!perm.isGranted) {
      setState(() => status = 'ظ„ط§ط²ظ… طھط³ظ…ط­ ط¨ط§ظ„ظˆطµظˆظ„ ظ„ظƒظ„ ط§ظ„ظ…ظ„ظپط§طھ ط§ظ„ط£ظˆظ„.');
      return;
    }
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
    if (!perm.isGranted) {
      setState(() => status = 'ظ„ط§ط²ظ… طھط³ظ…ط­ ط¨ط§ظ„ظˆطµظˆظ„ ظ„ظƒظ„ ط§ظ„ظ…ظ„ظپط§طھ ط§ظ„ط£ظˆظ„.');
      return;
    }
    if (!mounted) return;
    final res = await Navigator.push<String>(
      context,
      MaterialPageRoute(builder: (_) => FolderPicker(start: destDir)),
    );
    if (res != null) {
      try {
        final t = File('$res/.write_test')..writeAsStringSync('x');
        t.deleteSync();
      } catch (_) {
        setState(() => status = 'ظ…ط´ ظ‚ط§ط¯ط± ط£ظƒطھط¨ ظپظٹ ط§ظ„ظپظˆظ„ط¯ط± ط¯ظ‡: $res');
        return;
      }
      saveMOut(res);
      setState(() => destDir = res);
    }
  }

  
  Future<void> pickManualPdfs() async {
    try {
      final result = await FilePicker.pickFiles(type: FileType.custom, allowedExtensions: ['pdf']);
      if (result.isNotEmpty) {
        final paths = result.map((f) => f.path).whereType<String>().toList();
        _run(false, false, manualPaths: paths);
      }
    } catch (e) {
      _dbg('pickManualPdfs error: $e');
    }
  }

  Future<void> openFile(String? path, String mime) async {
    if (path == null || !File(path).existsSync()) {
      ScaffoldMessenger.of(context).showSnackBar(const SnackBar(content: Text('ط§ظ„ظ…ظ„ظپ ظ…ط´ ظ…ظˆط¬ظˆط¯.')));
      return;
    }
    try {
      await _chan.invokeMethod<bool>('openFile', {'path': path, 'mime': mime});
    } catch (_) {}
  }

  void useDefaultOut() {
    saveMOut(null);
    setState(() => destDir = defaultMDest);
  }

  Future<void> openDest() async {
    try {
      Directory(destDir).createSync(recursive: true);
    } catch (_) {}
    try {
      if (await chooseAndOpenFolder(context, destDir)) return;
    } catch (_) {}
    await Clipboard.setData(ClipboardData(text: destDir));
    if (!mounted) return;
    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(content: Text('ظ…ظ‚ط¯ط±طھط´ ط£ظپطھط­ ط§ظ„ظپظˆظ„ط¯ط±. ط§ظ„ظ…ط³ط§ط± ط§طھظ†ط³ط®: $destDir')),
    );
  }

  // â”€â”€â”€â”€â”€ ط§ظ„ظ…ط´ط§ط±ظƒط©: ظƒظ„ ط§ظ„ظ…ظ„ظپط§طھ ظ…ط±ط© ظˆط§ط­ط¯ط©طŒ ط£ظˆ ظ…ظ„ظپ ظ…ظ„ظپ â”€â”€â”€â”€â”€
  Future<void> shareAll() async {
    final files = results
        .where((r) => !r.excluded && r.newName != null && r.openPath != null && File(r.openPath!).existsSync())
        .map((r) => XFile(r.openPath!, mimeType: 'application/pdf'))
        .toList();
    if (files.isEmpty) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('ظ…ظپظٹط´ ظ…ظ„ظپط§طھ ظ…ط³ظ…ظ‘ط§ط© ظ„ظ…ط´ط§ط±ظƒطھظ‡ط§.')),
      );
      return;
    }
    if (!await confirmSend(context,
        review: results.where((r) => !r.ok).length,
        missing: missingMachines().length,
        pending: results.where((r) => r.pending).length)) {
      return;
    }
    await Share.shareXFiles(files, text: 'ط¥ظٹطµط§ظ„ط§طھ ط§ظ„ط¯ظپط¹ (${files.length} ظ…ظ„ظپ)');
  }

  // طµط­ = ظٹظپط¶ظ„ ظ…ط¹ ظ…ظ„ظپط§طھ ط§ظ„ظٹظˆظ…طŒ ط؛ظ„ط· = ظٹطھط´ط§ظ„ ظ…ظ† ط§ظ„ط¥ط±ط³ط§ظ„ ظˆط§ظ„ط¹ط¯ ظˆط§ظ„ظپظˆظ„ط¯ط±. ط§ظ„ط¶ط؛ط· طھط§ظ†ظٹ ط¨ظٹظ„ط؛ظٹ ط§ظ„ط§ط®طھظٹط§ط±
  void setVerdict(MResult r, int v, {bool toggle = true}) {
    final nv = (toggle && r.verdict == v) ? 0 : v;
    final wasExcluded = r.excluded;
    r.verdict = nv;
    final op = r.openPath;
    try {
      if (nv < 0 && !wasExcluded) {
        // ظ…ط³ط­ ط§ظ„ظ†ط³ط®ط© ط§ظ„ظ…طھط³ظ…ظٹط© ط¨ط³ (ظ…ط´ ط§ظ„ط£طµظ„ظٹ ط£ط¨ط¯ط§ظ‹)
        if (op != null && op != r.srcPath && op.startsWith(destDir)) File(op).deleteSync();
      } else if (nv >= 0 && wasExcluded && r.newName != null) {
        final src = File(r.srcPath);
        if (op != null && op != r.srcPath && src.existsSync()) {
          if (File(op).existsSync()) {
            r.openPath = r.srcPath; // ظپظٹظ‡ ظ…ظ„ظپ طھط§ظ†ظٹ ظˆط§ط®ط¯ ط§ظ„ط§ط³ظ… ط¯ظ‡
          } else {
            Directory(File(op).parent.path).createSync(recursive: true);
            src.copySync(op);
          }
        }
      }
    } catch (_) {}
    saveMState(results);
    setState(() {});
    // ط¹ظ„ظ‘ظ…طھظ‡ طµط­ ط¨ط³ ط§ظ„طھط·ط¨ظٹظ‚ ظ…ط´ ط¹ط§ط±ظپ ظ…ط§ظƒظٹظ†طھظ‡: ط§ط³ط£ظ„ ط¹ظ„ظٹظ‡ط§ ظˆط³ظ…ظ‘ظٹظ‡ ظˆط¹ط¯ظ‘ظ‡
    if (nv > 0 && r.newName == null && r.tidUsed == null && !busy) assignMachine(r);
  }

  // طھط­ط¯ظٹط¯ ط§ظ„ظ…ط§ظƒظٹظ†ط© ط¨ط¥ظٹط¯ظƒ ظ„ظ…ظ„ظپ ظ…ط´ ظ…ط¹ط±ظˆظپ: ط¨ظٹطھط³ظ…ظ‘ظ‰ ظˆط¨ظٹطھط¹ط¯ظ‘ ظˆط¨ظٹط¯ط®ظ„ ظپظٹ ط¥ظٹطµط§ظ„ط§طھ ط§ظ„ظٹظˆظ…
  Future<void> assignMachine(MResult r) async {
    if (busy) return;
    int? guess = r.orderNum;
    guess ??= int.tryParse(RegExp(r'^\s*(\d{1,2})').firstMatch(r.pdfName)?.group(1) ?? '');
    if (guess != null && !numToTid.containsKey(guess)) guess = null;
    int? sel = guess;
    final picked = await showDialog<int>(
      context: context,
      builder: (ctx) => StatefulBuilder(
        builder: (ctx, setS) => AlertDialog(
          title: const Text('ط¯ظ‡ ط¥ظٹطµط§ظ„ ط£ظ†ظ‡ظٹ ظ…ط§ظƒظٹظ†ط©طں'),
          content: DropdownButton<int>(
            isExpanded: true,
            value: sel,
            hint: const Text('ط§ط®طھط§ط± ط§ظ„ظ…ط§ظƒظٹظ†ط©'),
            items: [
              for (final i in machineNums())
                DropdownMenuItem(
                  value: i,
                  child: Text(machineLabel(i)),
                ),
            ],
            onChanged: (v) => setS(() => sel = v),
          ),
          actions: [
            TextButton(onPressed: () => Navigator.pop(ctx), child: const Text('ط¥ظ„ط؛ط§ط،')),
            FilledButton(
              onPressed: sel == null ? null : () => Navigator.pop(ctx, sel),
              child: const Text('طھظ…ط§ظ…'),
            ),
          ],
        ),
      ),
    );
    if (picked == null) return;
    final tid = numToTid[picked];
    String? day = r.dayUsed, month = r.monthUsed;
    if (day == null || month == null) {
      // ط§ظ„طھط§ط±ظٹط® ظ…ظ† ط§ط³ظ… ط§ظ„ظ…ظ„ظپ (ط¢ط®ط± ط´ظƒظ„ 7.11)طŒ ظˆط¥ظ„ط§ طھط§ط±ظٹط® ط§ظ„ظ†ظ‡ط§ط±ط¯ط©
      final ms = RegExp(r'(\d{1,2})\.(\d{1,2})').allMatches(r.pdfName).toList();
      for (final m in ms.reversed) {
        final d = int.parse(m.group(1)!), mo = int.parse(m.group(2)!);
        if (d >= 1 && d <= 31 && mo >= 1 && mo <= 12) {
          day = '$d';
          month = '$mo';
          break;
        }
      }
      day ??= '${DateTime.now().day}';
      month ??= '${DateTime.now().month}';
    }
    final srcF = File(r.srcPath).existsSync() ? File(r.srcPath) : File(r.openPath ?? r.srcPath);
    try {
      await _applyName(r, srcF, tid, day, month);
    } catch (e) {
      r.idOk = false;
      r.idNotes.add('ظپط´ظ„ ط§ظ„طھط³ظ…ظٹط©: ${e.toString().split('\n').first}');
    }
    r.verdict = 1;
    r.cntNotes.clear();
    r.count = -1; // ظٹطھط¹ط¯ظ‘ ظ…ظ† ط¬ط¯ظٹط¯
    r.tries = 0;
    r.detail = '';
    saveMState(results);
    if (!mounted) return;
    setState(() => status = 'ط§طھط³ظ…ظ‘ظ‰ ظ…ط§ظƒظٹظ†ط© $picked. ط§ط¶ط؛ط· \"ط¨ط¯ط، ط§ظ„ط¹ط¯\" ظ„ط¹ط¯ظ‘ ط¥ظٹطµط§ظ„ط§طھظ‡.');
  }

  Future<void> markRestOk() async {
    final list = results.where((r) => !r.ok).toList();
    if (list.isEmpty) return;
    final yes = await showDialog<bool>(
      context: context,
      builder: (ctx) => AlertDialog(
        title: const Text('ط§ظ„ط¨ط§ظ‚ظٹ طµط­'),
        content: Text('ظ‡طھط¹ظ„ظ‘ظ… ${list.length} ظ…ظ„ظپ ط¹ظ„ظ‰ ط¥ظ†ظ‡ظ… طµط­طں'),
        actions: [
          TextButton(onPressed: () => Navigator.pop(ctx, false), child: const Text('ظ„ط§')),
          FilledButton(onPressed: () => Navigator.pop(ctx, true), child: const Text('طھظ…ط§ظ…')),
        ],
      ),
    );
    if (yes != true) return;
    for (final r in list) {
      r.verdict = 1;
    }
    saveMState(results);
    if (mounted) setState(() {});
  }

  Future<void> shareOne(MResult r) async {
    final p = r.openPath;
    if (p == null || !File(p).existsSync()) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('ط§ظ„ظ…ظ„ظپ ظ…ط´ ظ…ظˆط¬ظˆط¯.')),
      );
      return;
    }
    await Share.shareXFiles([XFile(p, mimeType: 'application/pdf')],
        text: r.newName ?? r.pdfName);
  }

  Future<void> clearAll() async {
    if (results.isEmpty) {
      setState(() => status = 'ظ…ظپظٹط´ ط­ط§ط¬ط© ظ„ظ„ظ…ط³ط­.');
      return;
    }
    final mine = results.where((r) => r.newName != null && r.openPath != null).toList();
    final ok = await showDialog<bool>(
      context: context,
      builder: (ctx) => AlertDialog(
        title: const Text('ظ…ط³ط­ ط§ظ„ظƒظ„'),
        content: Text('ظ‡ظٹطھظ…ط³ط­ ${mine.length} ظ…ظ„ظپ ظ…ط³ظ…ظ‘ظ‰ ظ…ظ†:\n$destDir\n\n'
            'ظˆظ‡طھطھظ…ط³ط­ ظ†طھظٹط¬ط© ط§ظ„ظپط­طµ. (ظ…ظ„ظپط§طھ ظˆط§طھط³ط§ط¨ ط§ظ„ط£طµظ„ظٹط© ظ…ط´ ظ‡طھطھظ…ط³)'),
        actions: [
          TextButton(onPressed: () => Navigator.pop(ctx, false), child: const Text('ط¥ظ„ط؛ط§ط،')),
          FilledButton(onPressed: () => Navigator.pop(ctx, true), child: const Text('ظ…ط³ط­')),
        ],
      ),
    );
    if (ok != true) return;
    var n = 0;
    for (final r in mine) {
      try {
        if (r.openPath!.startsWith(destDir) || r.openPath!.startsWith(defaultMDest)) {
          File(r.openPath!).deleteSync();
          n++;
        }
      } catch (_) {}
    }
    for (final r in results) {
      r.ctl.dispose();
    }
    results.clear();
    saveMState(results);
    setState(() => status = 'طھظ… ظ…ط³ط­ $n ظ…ظ„ظپ ظˆظ†طھظٹط¬ط© ط§ظ„ظپط­طµ.');
  }

  // â”€â”€â”€â”€â”€ ط§ظ„طھط´ط؛ظٹظ„ â”€â”€â”€â”€â”€
  Future<void> start({bool resumeOnly = false, bool count = false}) async {
    if (busy) return;
    final perm = await Permission.manageExternalStorage.request();
    if (!perm.isGranted) {
      setState(() => status = 'ظ„ط§ط²ظ… طھط³ظ…ط­ ط¨ط§ظ„ظˆطµظˆظ„ ظ„ظƒظ„ ط§ظ„ظ…ظ„ظپط§طھ ط§ظ„ط£ظˆظ„.');
      return;
    }
    try {
      await Permission.notification.request();
    } catch (_) {}
    try {
      await _chan.invokeMethod(
          'startKeepAlive', {'title': 'A2R POS', 'text': 'ط¬ط§ط±ظٹ ظپط­طµ ط¥ظٹطµط§ظ„ط§طھ ط§ظ„ط¯ظپط¹...'});
    } catch (_) {}
    try {
      await _chan.invokeMethod('keepScreenOn', {'on': true});
    } catch (_) {}
    try {
      await _run(resumeOnly, count);
    } finally {
      try {
        await _chan.invokeMethod('stopKeepAlive');
      } catch (_) {}
      try {
        await _chan.invokeMethod('keepScreenOn', {'on': false});
      } catch (_) {}
      if (_life != AppLifecycleState.resumed) {
        final total = results.fold<int>(0, (a, r) => a + (r.count > 0 && !r.excluded ? r.count : 0));
        try {
          await _chan.invokeMethod('notify', {
            'title': tr('A2R POS - ط§ظ†طھظ‡ظ‰ ظپط­طµ ط¥ظٹطµط§ظ„ط§طھ ط§ظ„ط¯ظپط¹'),
            'text': tr('${results.length} ظ…ظ„ظپطŒ $total ط¥ظٹطµط§ظ„')
          });
        } catch (_) {}
      }
    }
  }

  Future<void> _run(bool resumeOnly, bool count, {List<String>? manualPaths}) async {
    await ensureDefaultFiles();
    await loadSettingsFromExcel();
    setState(() {
      busy = true;
      stopReq = false;
      progressValue = 0;
      status = 'ط¬ط§ط±ظٹ ط§ظ„ط¨ط­ط« ط¹ظ† ظ…ظ„ظپط§طھ PDF...';
    });
    try {
      Directory(destDir).createSync(recursive: true);
      Directory(tmpDir).createSync(recursive: true);
    } catch (_) {}
    final rec = TextRecognizer(script: TextRecognitionScript.latin);
    String? idle;
    try {
      // ط§ظ„ظ…ط±ط­ظ„ط© 1 (ط³ط±ظٹط¹ط©): طھط¹ط±ظٹظپ ط§ظ„ظˆط­ط¯ط© ظˆطھط³ظ…ظٹط© ظƒظ„ ط§ظ„ظ…ظ„ظپط§طھ
      try {
        final lf = File('$baseDir/id_debug.txt');
        if (!resumeOnly && lf.existsSync() && lf.lengthSync() > 300000) lf.deleteSync();
      } catch (_) {}
      _dbg('PHASE1 start resumeOnly=$resumeOnly');
      namingDone = resumeOnly;
      if (!resumeOnly) idle = await _identifyAll(rec, manualPaths: manualPaths);
      _dbg('PHASE1 done');
      if (mounted) setState(() => namingDone = true);
      // ط§ظ„ظ…ط±ط­ظ„ط© 2 (ط¨ط·ظٹط¦ط©): ط¹ط¯ظ‘ ط§ظ„ط¥ظٹطµط§ظ„ط§طھ ط¬ظˆظ‡ ظƒظ„ ظ…ظ„ظپ: ظ…ط´ ط¨طھط´طھط؛ظ„ ط؛ظٹط± ظ„ظ…ط§ طھط¶ط؛ط· "ط¨ط¯ط، ط§ظ„ط¹ط¯"
      if (count) {
        _dbg('PHASE2 start');
        await _countAll(rec);
        _dbg('PHASE2 done');
      }
      _flagDuplicates();
    } finally {
      await rec.close();
      try {
        Directory(tmpDir).deleteSync(recursive: true);
      } catch (_) {}
    }
    results.sort(_cmpM);
    saveMState(results);
    final pend = results.where((r) => r.pending).length;
    final total = results.fold<int>(0, (a, r) => a + (r.count > 0 && !r.excluded ? r.count : 0));
    final named = results.where((r) => r.newName != null && !r.excluded).length;
    if (!mounted) return;
    setState(() {
      busy = false;
      progressValue = 0;
      if (stopReq) {
        status = 'ط§طھظˆظ‚ظپ. ظپط§ط¶ظ„ $pend ظ…ظ„ظپ ظ„ط³ظ‡ ظ…ط§ ط§طھط¹ط¯ظ‘ط´طŒ ط§ط¶ط؛ط· "ظپط­طµ" ظˆظ‡ظٹظƒظ…ظ‘ظ„ ظ…ظ† ظ…ظƒط§ظ†ظ‡.';
      } else if (results.isEmpty && idle != null) {
        status = idle;
      } else if (!count) {
        status = 'طھظ…طھ ط§ظ„طھط³ظ…ظٹط©: ${results.length} ظ…ظ„ظپطŒ $named ط§طھط³ظ…ظ‘ظ‰ ظˆط§طھظ†ط³ط®.'
            '${pend > 0 ? ' ط§ط¶ط؛ط· \"ط¨ط¯ط، ط§ظ„ط¹ط¯\" ظ„ط¹ط¯ظ‘ ط¥ظٹطµط§ظ„ط§طھ $pend ظ…ظ„ظپ.' : ''}';
      } else {
        status = 'ط§ظƒطھظ…ظ„ ط§ظ„ط¹ط¯: ${results.length} ظ…ظ„ظپطŒ $named ط§طھط³ظ…ظ‘ظ‰طŒ ط¥ط¬ظ…ط§ظ„ظٹ $total ط¥ظٹطµط§ظ„.'
            '${pend > 0 ? ' (ظپط§ط¶ظ„ $pend ظ„ط³ظ‡)' : ''}';
      }
    });
  }

  // ط¨ظٹط±ط¬ظ‘ط¹ ط±ط³ط§ظ„ط© ظ„ظˆ ظ…ظپظٹط´ ط­ط§ط¬ط© ط¬ط¯ظٹط¯ط© ظٹظپط­طµظ‡ط§
  Future<String?> _identifyAll(TextRecognizer rec, {List<String>? manualPaths}) async {
    final dirs = pdfDir != null ? [Directory(pdfDir!)] : findWhatsAppDocDirs();
    if (manualPaths != null) dirs.clear();
    final now = DateTime.now();
    final cutoff = days < 0
        ? DateTime(2000)
        : DateTime(now.year, now.month, now.day).subtract(Duration(days: days));
    final have = results.map((r) => r.key).toSet();
    _dbg('DIRS: ${dirs.map((d) => d.path).join(' | ')}');
    final pdfs = <File>[];
    var skipped = 0;
    for (final d in dirs) {
      if (d.path == destDir || d.path == defaultMDest) continue; // ظ†ط§طھط¬ ط§ظ„طھط·ط¨ظٹظ‚ ظ†ظپط³ظ‡
      try {
        for (final e in d.listSync().whereType<File>()) {
          final n = baseName(e.path);
          final fd = fileDate(e);
          var isPdf = n.toLowerCase().trim().endsWith('.pdf');
          if (!isPdf && !fd.isBefore(cutoff) && !n.substring(n.lastIndexOf('.') + 1).contains(RegExp(r'[A-Za-z0-9]'))) {
            isPdf = _hasPdfHeader(e); // ظ…ظ„ظپط§طھ ظˆط§طھط³ط§ط¨ DOC-... ظ…ظ† ط؛ظٹط± ط§ظ…طھط¯ط§ط¯
          }
          if (!isPdf) {
            if (!fd.isBefore(cutoff)) _dbg('SKIP(ظ…ط´ PDF) $n  size=${e.lengthSync()}');
            continue;
          }
          // ط¥ط®ط±ط§ط¬ ط§ظ„طھظ‚ظپظٹظ„ ط¨طھط§ط¹ ط§ظ„طھط·ط¨ظٹظ‚ ط¨ط³ (ظ†طھظٹط¬ط© 5.11.pdf)طŒ ط£ظٹ ط§ط³ظ… طھط§ظ†ظٹ ط¨ظٹطھظپط­طµ ط¹ط§ط¯ظٹ
          if (_isAppReportName(n)) {
            _dbg('SKIP(طھظ‚ط±ظٹط± ط§ظ„طھط·ط¨ظٹظ‚) $n');
            continue;
          }
          if (fd.isBefore(cutoff)) {
            if (fd.isAfter(cutoff.subtract(const Duration(days: 14)))) {
              _dbg('SKIP(ط£ظ‚ط¯ظ… ظ…ظ† ط§ظ„ظپطھط±ط©: $fd) $n');
            }
            continue;
          }
          if (have.contains('$n:${e.lengthSync()}') && skipDone) {
            skipped++;
            _dbg('SKIP(ط§طھظپط­طµ ظ‚ط¨ظ„ ظƒط¯ظ‡) $n');
            continue;
          }
          pdfs.add(e);
        }
      } catch (_) {}
    }
    pdfs.sort((a, b) => baseName(a.path).compareTo(baseName(b.path)));
    if (manualPaths != null) pdfs.addAll(manualPaths.map((p) => File(p)));
    if (pdfs.isEmpty) {
      if (dirs.isEmpty) return 'ظ…ط´ ظ„ط§ظ‚ظٹ ظپظˆظ„ط¯ط± ظ…ط³طھظ†ط¯ط§طھ ظˆط§طھط³ط§ط¨. ط§ط®طھط§ط± ظپظˆظ„ط¯ط± ط¨ط¥ظٹط¯ظƒ.';
      return skipped > 0
          ? 'ظƒظ„ ط§ظ„ظ…ظ„ظپط§طھ ظپظٹ ط§ظ„ظپطھط±ط© ط¯ظٹ ط§طھظپط­طµطھ ظ‚ط¨ظ„ ظƒط¯ظ‡ ($skipped ظ…ظ„ظپ).'
          : 'ظ…ظپظٹط´ ظ…ظ„ظپط§طھ PDF ظپظٹ ط§ظ„ظپطھط±ط© ط¯ظٹ.';
    }
    // ط§ظ…ط³ط­ ظƒظ„ ط§ظ„ظ†طھط§ظٹط¬/ط§ظ„ظ†ط³ط® ط§ظ„ظ‚ط¯ظٹظ…ط© ظ„ظ„ظ…ظ„ظپط§طھ ط§ظ„ظ„ظٹ ظ‡ظ†ظپط­طµظ‡ط§ ط¯ظ„ظˆظ‚طھظٹ *ظ‚ط¨ظ„* ط§ظ„طھط³ظ…ظٹط©طŒ
    // ط¹ظ„ط´ط§ظ† ط£ط³ظ…ط§ط، ظ‚ط¯ظٹظ…ط© ظ…طھط¹ظ…ظ„ط´ طھط¶ط§ط±ط¨ ظˆطھط·ظ„ظ‘ط¹ "_2" ظ…ظ† ط؛ظٹط± ظ…ط§ ظٹظƒظˆظ† ظپظٹظ‡ ظ…ظ„ظپ ظ…ظƒط±ط±.
    for (final pf in pdfs) {
      final pn = baseName(pf.path);
      final ps = pf.lengthSync();
      for (final o in results.where((x) => x.key == '$pn:$ps' || x.srcPath == pf.path).toList()) {
        try {
          if (o.newName != null && o.openPath != null && o.openPath != o.srcPath) {
            File(o.openPath!).deleteSync();
          }
        } catch (_) {}
        o.ctl.dispose();
        results.remove(o);
      }
    }
    for (var i = 0; i < pdfs.length; i++) {
      if (stopReq) break;
      final f = pdfs[i];
      final name = baseName(f.path);
      final size = f.lengthSync();
      // ظ„ظˆ ط§ظ„ظ…ظ„ظپ ط¯ظ‡ ط§طھظپط­طµ ظ‚ط¨ظ„ ظƒط¯ظ‡ ظˆط¥ط­ظ†ط§ ط¨ظ†ط¹ظٹط¯: ط§ظ…ط³ط­ ط§ظ„ط¥ط¯ط®ط§ظ„ ط§ظ„ظ‚ط¯ظٹظ… ظˆظ†ط³ط®طھظ‡
      for (final o in results.where((x) => x.key == '$name:$size').toList()) {
        try {
          if (o.newName != null && o.openPath != null) File(o.openPath!).deleteSync();
        } catch (_) {}
        o.ctl.dispose();
        results.remove(o);
      }
      final r = MResult(name)
        ..srcPath = f.path
        ..size = size;
      setState(() => status = 'طھط¹ط±ظٹظپ ط§ظ„ظ…ظ„ظپط§طھ ${i + 1} / ${pdfs.length}: $name');
      try {
        await _identifyOne(r, f, rec);
      } catch (e, st) {
        r.idOk = false;
        r.idNotes.add('ظپط´ظ„ ط§ظ„طھط¹ط±ظٹظپ: ${e.toString().split('\n').first}');
        final ss = st.toString();
        _dbg('ERROR(طھط¹ط±ظٹظپ) $name size=$size: $e\n${ss.substring(0, ss.length > 600 ? 600 : ss.length)}');
      }
      r.openPath ??= f.path;
      results.add(r);
      saveMState(results);
      _dbg('FILE ${i + 1}/${pdfs.length} done: $name  RSS=${ProcessInfo.currentRss ~/ 1048576}MB');
      if (!mounted) return null;
      setState(() => progressValue = (i + 1) / pdfs.length);
    }
    _dbg('resolveDuplicates start');
    await _resolveDuplicates();
    _flagDuplicates();
    results.sort(_cmpM);
    saveMState(results);
    _dbg('resolveDuplicates done');
    return null;
  }

  // ظ‚ط±ط§ط،ط© طµظپط­ط§طھ ط§ظ„ظ€ PDF (ظ„ط­ط¯ 6) ظˆطھطµظˆظٹطھ ط¹ظ„ظ‰ Terminal IDطŒ ظˆط¨ط¹ط¯ظٹظ† ظ…ط·ط§ط¨ظ‚طھظ‡ط§ ط¨ط§ط³ظ… ط§ظ„ظ…ظ„ظپ ط§ظ„ط£طµظ„ظٹ
  Future<void> _identifyOne(MResult r, File f, TextRecognizer rec) async {
    final bytes = await f.readAsBytes();
    final sb = StringBuffer();
    final votes = <String, double>{};
    var id = MIdent();
    for (var pi = 0; pi < 25; pi++) {
      // ط¨ط¹ط¯ 6 طµظپط­ط§طھ ظ„ظˆ ظ„ظ‚ظٹظ†ط§ ظ…ط§ظƒظٹظ†ط© ظ†ظˆظ‚ظپطŒ ظˆظ„ظˆ ظ„ط³ظ‡ ظ…ظپظٹط´ ظ†ظƒظ…ظ‘ظ„ ظ„ط­ط¯ 25 طµظپط­ط©
      if (pi >= 6 && votes.isNotEmpty) break;
      var got = false;
      try {
        final dpi = await _safeDpi(bytes, pi, 200);
        await for (final page in Printing.raster(bytes, pages: [pi], dpi: dpi)) {
          got = true;
          final png = await page.toPng();
          final tp = File('$tmpDir/id_$pi.png');
          await tp.writeAsBytes(png);
          try {
            final res = await rec.processImage(InputImage.fromFilePath(tp.path));
            sb.writeln(res.text);
            _tidVotes(res.text).forEach((k, v) => votes[k] = (votes[k] ?? 0) + v);
          } catch (_) {}
          try {
            tp.deleteSync();
          } catch (_) {}
        }
      } catch (e) {
        if (!_pageEnd(e) || pi == 0) rethrow; // ط¢ط®ط± طµظپط­ط© ط¹ط¯ظ‘طھ = ظ…ط´ ط®ط·ط£
      }
      if (!got) break;
      id = _identify(sb.toString());
      final rk = _rankVotes(votes);
      // ظˆظ‚ظپ ط¨ط¯ط±ظٹ ظ„ظˆ ط§ظ„ظ…ط§ظƒظٹظ†ط© ظˆط§ط¶ط­ط© (Terminal ID ظ…ظ‚ط±ظˆط، ط¨ظˆط¶ظˆط­) ظˆط§ظ„ظٹظˆظ… ظ…ط¹ط±ظˆظپ
      if (id.day != null &&
          rk.isNotEmpty &&
          rk.first.value >= 3 &&
          (rk.length < 2 || rk.first.value >= rk[1].value * 2)) {
        break;
      }
    }
    final rk = _rankVotes(votes);
    r.altTids = rk.map((e) => e.key).toList();
    r.votes = votes;
    var tid = rk.isNotEmpty ? rk.first.key : id.tid;
    var fromUnit = false;
    if (tid == null) {
      // ظ…ظپظٹط´ Terminal ID ظ…ظ‚ط±ظˆط،: ظ†ط­ط¯ط¯ ط§ظ„ظˆط­ط¯ط© ظ…ظ† Merchant ID/ط§ط³ظ… ط§ظ„ظˆط­ط¯ط© ط¬ظˆظ‡ ط§ظ„ظ€ PDFطŒ
      // ظˆظ„ظˆ ط§ظ„ظˆط­ط¯ط© ظپظٹظ‡ط§ ظ…ط§ظƒظٹظ†ط© ظˆط§ط­ط¯ط© ط¨ط³ ظ†ط³طھط®ط¯ظ…ظ‡ط§ (ظˆطھطھط¹ظ„ظ‘ظ… ظ„ظ„ظ…ط±ط§ط¬ط¹ط©)
      final row = findUnitExtraIn([sb.toString()]);
      if (row != null) {
        final same = terminals.entries.where((e) => e.value.excelRow == row).toList();
        if (same.length == 1) {
          tid = same.first.key;
          fromUnit = true;
        }
      }
    }
    try {
      final all = sb.toString();
      File('$baseDir/id_debug.txt').writeAsStringSync(
          '##### ${r.pdfName}\ntid=$tid fromUnit=$fromUnit day=${id.day}.${id.month}\n'
          'votes=${rk.map((e) => '${e.key}:${e.value}').join(' , ')}\n'
          '${tid == null ? all.substring(0, all.length > 2500 ? 2500 : all.length) : ''}\n\n',
          mode: FileMode.append);
    } catch (_) {}
    final top = rk.isNotEmpty ? rk.first.value : 0.0;
    final weak = tid != null &&
        ((!fromUnit && top < 2) || !_plausibleDay(id.day, id.month));
    if (weak) {
      // ظ…ط´ ظˆط§ط¶ط­ ط¥ظ†ظ‡ ط¥ظٹطµط§ظ„ ظ…ط§ظƒظٹظ†ط© (طھظˆظƒظٹظ„ / ظ…ظ„ظپ طھط§ظ†ظٹ): طھط­طھ ط§ظ„ظ…ط±ط§ط¬ط¹ط©طŒ ظ…ظ† ط؛ظٹط± طھط³ظ…ظٹط© ظˆظ„ط§ ط¹ط¯
      r.tidUsed = null;
      r.newName = null;
      if (id.day != null && id.month != null) {
        r.dayUsed = id.day;
        r.monthUsed = id.month;
      }
      r.idOk = false;
      r.idNotes.add('ط§ظ„ظ…ظ„ظپ ط¯ظ‡ ظ…ط´ ظˆط§ط¶ط­ ط¥ظ†ظ‡ ط¥ظٹطµط§ظ„ ظ…ط§ظƒظٹظ†ط© (ظ‚ط±ط§ط،ط© ط¶ط¹ظٹظپط© ظ„ظ„ظ€ Terminal ID)طŒ ط±ط§ط¬ط¹ظ‡');
      r.count = 0;
      r.cntOk = true;
      r.detail = 'ظ„ظ„ظ…ط±ط§ط¬ط¹ط© (ظ…ط´ ظ…ط¹ط¯ظˆط¯)';
      return;
    }
    await _applyName(r, f, tid, id.day, id.month);
    if (fromUnit && r.newName != null) {
      r.idOk = false;
      r.idNotes.add('Terminal ID ظ…ط´ ظˆط§ط¶ط­طŒ ط§ظ„ظˆط­ط¯ط© ط§طھط­ط¯ط¯طھ ظ…ظ† Merchant ID/ط§ظ„ط§ط³ظ… ط¬ظˆظ‡ ط§ظ„ظ€ PDFطŒ ط±ط§ط¬ط¹');
    }
  }

  // ظ„ظˆ ظ…ظ„ظپظٹظ† ط§طھط¹ط±ظپظˆط§ ط¹ظ„ظ‰ ظ†ظپط³ ط§ظ„ظ…ط§ظƒظٹظ†ط©: ط§ظ„ط£ظ‚ظˆظ‰ (ط£ظƒطھط± ط£طµظˆط§طھ) ظٹظپط¶ظ„طŒ ظˆط§ظ„طھط§ظ†ظٹ ظٹط§ط®ط¯ ظ…ط±ط´ط­ظ‡ ط§ظ„طھط§ظ†ظٹ ظ„ظˆ ظ‚ط±ط§ط،طھظ‡ ظˆط§ط¶ط­ط©
  Future<void> _resolveDuplicates() async {
    final groups = <String, List<MResult>>{};
    for (final r in results) {
      if (r.tidUsed != null && r.newName != null && !r.excluded) {
        groups.putIfAbsent(r.tidUsed!, () => <MResult>[]).add(r);
      }
    }
    for (final e in groups.entries) {
      final g = e.value;
      if (g.length < 2) continue;
      g.sort((a, b) => (b.votes[e.key] ?? 0).compareTo(a.votes[e.key] ?? 0));
      for (var i = 1; i < g.length; i++) {
        final r = g[i];
        if (!File(r.srcPath).existsSync()) continue;
        final claimed = results.map((x) => x.tidUsed).whereType<String>().toSet();
        String? alt;
        for (final c in r.altTids) {
          if (c == e.key || claimed.contains(c)) continue;
          if ((r.votes[c] ?? 0) >= 1) {
            alt = c;
            break;
          }
        }
        if (alt == null) continue;
        await _applyName(r, File(r.srcPath), alt, r.dayUsed, r.monthUsed);
        if (r.newName != null) {
          r.idOk = false;
          r.idNotes.add('ط§ظ„ظ€ PDF ط§طھط¹ط±ظ‘ظپ ط¹ظ„ظ‰ ظ†ظپط³ ظ…ط§ظƒظٹظ†ط© ظ…ظ„ظپ طھط§ظ†ظٹطŒ ط§طھط³ظ…ظ‘ظ‰ ظ…ط§ظƒظٹظ†ط© ${terminals[alt]!.orderNum} (ط§ظ„ظ…ط±ط´ط­ ط§ظ„طھط§ظ†ظٹ ظ…ظ† ظ‚ط±ط§ط،ط© ط§ظ„ظ€ PDF)طŒ ط±ط§ط¬ط¹');
        }
      }
    }
  }

  // ط¨ظٹط¹ظ„ظ‘ظ… ط§ظ„ظ…ظ„ظپط§طھ ط§ظ„ظ„ظٹ ظ„ط³ظ‡ ظˆط§ط®ط¯ط© ظ†ظپط³ ط§ظ„ظ…ط§ظƒظٹظ†ط© (ظ…ظپظٹط´ ط­ظ„ طھظ„ظ‚ط§ط¦ظٹ)
  void _flagDuplicates() {
    const tag = 'ظ†ظپط³ ط§ظ„ظ…ط§ظƒظٹظ†ط©';
    for (final r in results) {
      final had = r.idNotes.any((n) => n.startsWith(tag));
      r.idNotes.removeWhere((n) => n.startsWith(tag));
      if (had && r.idNotes.isEmpty && r.newName != null) r.idOk = true;
    }
    final groups = <String, List<MResult>>{};
    for (final r in results) {
      if (r.tidUsed != null && r.newName != null && !r.excluded) {
        groups.putIfAbsent(r.tidUsed!, () => <MResult>[]).add(r);
      }
    }
    for (final e in groups.entries) {
      if (e.value.length < 2) continue;
      final info = terminals[e.key];
      for (final r in e.value) {
        final others = e.value.where((x) => !identical(x, r)).map((x) => x.pdfName).join(' طŒ ');
        r.idOk = false;
        r.idNotes.add('$tag (${info?.orderNum ?? ''} ${info?.unitName ?? ''}) ط§طھط¹ط±ظپطھ ظƒظ…ط§ظ† ظپظٹ: $others â€” ط±ط§ط¬ط¹');
      }
    }
  }

  Future<void> _applyName(MResult r, File f, String? tid, String? day, String? month) async {
    r.idOk = true;
    r.idNotes.clear();
    // ظ†ط³ط®ط© ظ‚ط¯ظٹظ…ط© ظ…طھط³ظ…ظٹط© ظ„ظ„ظ…ظ„ظپ ط¯ظ‡: ظ†ظ…ط³ط­ظ‡ط§ ظ‚ط¨ظ„ ظ…ط§ ظ†ط³ظ…ظ‘ظٹ ظ…ظ† ط¬ط¯ظٹط¯
    if (r.newName != null &&
        r.openPath != null &&
        r.openPath != r.srcPath &&
        r.openPath != f.path &&
        r.openPath!.startsWith(destDir)) {
      try {
        File(r.openPath!).deleteSync();
      } catch (_) {}
    }
    r.newName = null;
    r.tidUsed = tid;
    if (tid == null || !terminals.containsKey(tid)) {
      r.idOk = false;
      r.idNotes.add('ظ…ط´ ظ‚ط§ط¯ط± ط£ط¹ط±ظپ ط§ظ„ظ…ط§ظƒظٹظ†ط© ظ…ظ† ط§ظ„ظ€ PDF');
      return;
    }
    final info = terminals[tid]!;
    r.orderNum = info.orderNum;
    if (day == null || month == null) {
      r.idOk = false;
      r.idNotes.add('ظ…ط´ ظ‚ط§ط¯ط± ط£ط¹ط±ظپ ط§ظ„طھط§ط±ظٹط® ظ…ظ† ط§ظ„ظ€ PDF');
      return;
    }
    r.dayUsed = day;
    r.monthUsed = month;
    final name = machineNewName(info, day, month);
    var final0 = name;
    // "_2" ط¨ط³ ظ„ظˆ ظپط¹ظ„ط§ظ‹ ظ…ظ„ظپ طھط§ظ†ظٹ (ظ…ظ† ظ†ظپط³ ط§ظ„ظپط­طµ) ظˆط§ط®ط¯ ظ†ظپط³ ط§ظ„ط§ط³ظ…ط› ط£ظٹ ظ…ظ„ظپ ظ‚ط¯ظٹظ… ظپظٹ ط§ظ„ظپظˆظ„ط¯ط± ط¨ظ†ظƒطھط¨ ظپظˆظ‚ظ‡
    var k = 2;
    while (results.any((x) => x != r && !x.excluded && x.newName == final0)) {
      final0 = '${name}_$k';
      k++;
    }
    Directory(destDir).createSync(recursive: true);
    final dst = File('$destDir/$final0.pdf');
    if (dst.path != f.path) {
      try {
        if (dst.existsSync()) dst.deleteSync();
      } catch (_) {}
      await f.copy(dst.path);
    }
    r.newName = final0;
    r.openPath = dst.path;
  }

  Future<void> _countAll(TextRecognizer rec) async {
    // ظ…ظ„ظپ ط§ظ„طھط·ط¨ظٹظ‚ ط§طھظ‚ظپظ„ ط£ط«ظ†ط§ط، ط¹ط¯ظ‘ظ‡ 3 ظ…ط±ط§طھ: ظ†ط¹ط¯ظ‘ظٹظ‡ ط¨ط¯ظ„ ظ…ط§ ظ†ظ„ظپ ظپظٹ ط¯ط§ظٹط±ط©
    for (final r in results) {
      if (r.pending && r.tries >= 3) {
        r.count = 0;
        r.cntOk = false;
        r.cntNotes.add('ط§ظ„ط¹ط¯ ظˆظ‚ظپ ط£ظƒطھط± ظ…ظ† ظ…ط±ط© (ط§ظ„طھط·ط¨ظٹظ‚ ط§طھظ‚ظپظ„). ظ…ظ† â‹® ط§ط®طھط§ط± ط¥ط¹ط§ط¯ط© ط§ظ„ط¹ط¯');
      }
    }
    final todo = results.where((r) => r.pending).toList();
    for (var i = 0; i < todo.length; i++) {
      if (stopReq) break;
      final r = todo[i];
      if (r.excluded) continue; // ط§طھط¹ظ„ظ‘ظ… ط؛ظ„ط· ظˆط¥ط­ظ†ط§ ط¨ظ†ط¹ط¯
      r.tries++;
      saveMState(results);
      setState(() {
        status = 'ط¹ط¯ظ‘ ط§ظ„ط¥ظٹطµط§ظ„ط§طھ ${i + 1} / ${todo.length}: ${r.newName ?? r.pdfName}';
        progressValue = i / todo.length;
      });
      try {
        await _countOne(r, rec);
      } catch (_) {
        r.count = 0;
        r.cntOk = false;
        r.cntNotes.add('ظپط´ظ„ ط§ظ„ط¹ط¯');
      }
      r.tries = 0;
      saveMState(results);
      _dbg('COUNT ${i + 1}/${todo.length} ${r.newName ?? r.pdfName} = ${r.count}  RSS=${ProcessInfo.currentRss ~/ 1048576}MB');
      if (!mounted) return;
      setState(() => progressValue = (i + 1) / todo.length);
    }
  }


  // ط§ظ„ط¹ط¯ ط§ظ„ط³ط±ظٹط¹: ط¨ظٹظ‚ط±ط§ طµظˆط±ط© ظƒظ„ طµظپط­ط© ظ…ط±ط© ظˆط§ط­ط¯ط© ظ…ظ† ط؛ظٹط± طھط­ظˆظٹظ„ PNG ظˆظ„ط§ ط´ط±ط§ط¦ط­. ط¨ظٹط±ط¬ظ‘ط¹ false ظ„ظˆ ط§ظ„ظ…ظ„ظپ ظ…ط´ ظ…ظ…ط³ظˆط­
  Future<bool> _countOneFast(MResult r, TextRecognizer rec, Uint8List bytes, File src) async {
    final ranges = _pdfJpegRanges(bytes);
    if (ranges.isEmpty) return false;
    final col = MCollector();
    final sw = Stopwatch()..start();
    var total = 0, totalMx = 0, failed = 0;
    final diff = <String>[];
    final empty = <int>[];
    for (var i = 0; i < ranges.length; i++) {
      final fp = File('$tmpDir/f_$i.jpg');
      try {
        await fp.writeAsBytes(Uint8List.sublistView(bytes, ranges[i][0], ranges[i][1]));
        final res = await rec.processImage(InputImage.fromFilePath(fp.path));
        final sp = _structCount(res);
        total += sp.sm;
        totalMx += sp.mx;
        if (sp.sm == 0) empty.add(i + 1);
        if ((sp.sm - sp.mx).abs() >= 2) diff.add('${i + 1} (${sp.sm} ط£ظˆ ${sp.mx})');
        col.pageH = sp.h;
        col.pages.add(<Anchor>[]);
        _collect(res, 0, 1, col);
      } catch (_) {
        failed++;
      }
      try {
        fp.deleteSync();
      } catch (_) {}
      await Future.delayed(const Duration(milliseconds: 30));
    }
    _dbg('FAST ${r.pdfName}: ${ranges.length} طµظپط­ط©طŒ ${sw.elapsedMilliseconds ~/ 1000}ط«طŒ ط¹ط¯ط¯ $total (ط¨ط§ظ„ط·ط±ظٹظ‚ط© ط§ظ„طھط§ظ†ظٹط© $totalMx)طŒ ظپط´ظ„ $failed');
    if (total == 0 || failed == ranges.length) return false; // ظ†ط±ط¬ط¹ ظ„ظ„ط·ط±ظٹظ‚ط© ط§ظ„ظ‚ط¯ظٹظ…ط©
    final sc = col.finish();
    r.count = total;
    r.detail = 'ط¹ط¯ظ‘ ط³ط±ظٹط¹ â€¢ ${ranges.length} طµظپط­ط© â€¢ RRN ظ…ظ‚ط±ظˆط، ${sc.full}${sc.rangeText}';
    if (failed > 0) {
      r.cntOk = false;
      r.cntNotes.add('$failed طµظپط­ط© ظپط´ظ„طھ ظ‚ط±ط§ط،طھظ‡ط§طŒ ط§ظ„ط¹ط¯ط¯ ظ…ظ…ظƒظ† ظٹظƒظˆظ† ظ†ط§ظ‚طµ');
    }
    if (empty.isNotEmpty) {
      r.cntNotes.add('طµظپط­ط§طھ ظ…ط§ ظ„ظ‚ظٹطھط´ ظپظٹظ‡ط§ ط¥ظٹطµط§ظ„ط§طھ: ${empty.join('طŒ ')}');
    }
    if (diff.isNotEmpty) {
      r.cntNotes.add('طµظپط­ط§طھ ط§ظ„ط·ط±ظٹظ‚طھظٹظ† ط§ط®طھظ„ظپظˆط§ ظپظٹظ‡ط§: ${diff.join('طŒ ')} (ط±ط§ط¬ط¹ظ‡ط§ ظ„ظˆ ط§ظ„ط¹ط¯ط¯ ظ…ظ‡ظ…)');
    }
    // ظ†ظ‚ط§ط±ظ† ط¨طھط³ظ„ط³ظ„ ط§ظ„ظ€ RRN ط¨ط³ ظ„ظˆ ط§ظ„ط£ط±ظ‚ط§ظ… ط§طھظ‚ط±طھ ظƒظپط§ظٹط©
    if (sc.rangeLen != null && sc.full >= total * 0.8 && total != sc.rangeLen) {
      r.cntOk = false;
      final miss = sc.seqMissing.take(10).map((n) => n.toString().padLeft(6, '0')).join('طŒ ');
      r.cntNotes.add('ط§ظ„ط¹ط¯ط¯ $total ط¨ط³ ط§ظ„طھط³ظ„ط³ظ„ ${sc.rangeLen}'
          '${sc.seqMissing.isNotEmpty ? ' - ط£ط±ظ‚ط§ظ… ظ…ط´ ظ„ط§ظ‚ظٹظ‡ط§: $miss${sc.seqMissing.length > 10 ? ' ...' : ''}' : ''}');
    }
    if (sc.otherDay.length >= 3 && sc.otherDay.length >= total * 0.1) {
      r.cntOk = false;
      r.cntNotes.add('ظ…ط´ طھط¨ط¹ ظ†ظپط³ ط§ظ„ظٹظˆظ…: ${sc.otherDay.take(5).join(' طŒ ')}');
    }
    if (sc.nonPayment) {
      r.cntNotes.add('ظپظٹظ‡ ط¥ظٹطµط§ظ„ ط¥ظ‚ظپط§ظ„/ظ…ظ„ط®طµ/ط¥ط¹ط§ط¯ط© ط·ط¨ط§ط¹ط© (ط§طھط¬ط§ظ‡ظ„ ظˆظ…ط´ ظ…ط¹ط¯ظˆط¯)');
    }
    if (r.newName == null && sc.tid != null) {
      await _applyName(r, src, sc.tid, sc.day, sc.month);
    } else if (sc.tid != null && terminals[sc.tid] != null && r.orderNum != null &&
        terminals[sc.tid]!.orderNum != r.orderNum) {
      r.cntOk = false;
      r.cntNotes.add('ط§ظ„ط¹ط¯ ظ„ظ‚ظ‰ ظ…ط§ظƒظٹظ†ط© ${terminals[sc.tid]!.orderNum} ط¨ط³ ط§ظ„طھط³ظ…ظٹط© ظ…ط§ظƒظٹظ†ط© ${r.orderNum}طŒ ط±ط§ط¬ط¹');
    }
    return true;
  }

  Future<void> _countOne(MResult r, TextRecognizer rec) async {
    File? src;
    if (File(r.srcPath).existsSync()) {
      src = File(r.srcPath);
    } else if (r.openPath != null && File(r.openPath!).existsSync()) {
      src = File(r.openPath!);
    }
    r.cntNotes.clear();
    r.cntOk = true;
    if (src == null) {
      r.count = 0;
      r.cntOk = false;
      r.cntNotes.add('ط§ظ„ظ…ظ„ظپ ط§ظ„ط£طµظ„ظٹ ظ…ط´ ظ…ظˆط¬ظˆط¯');
      return;
    }
    final bytes = await src.readAsBytes();
    if (fastCount) {
      try {
        if (await _countOneFast(r, rec, bytes, src)) return;
      } catch (e) {
        _dbg('FAST ظپط´ظ„ ظˆظ‡ظ†ظƒظ…ظ„ ط¨ط§ظ„ط·ط±ظٹظ‚ط© ط§ظ„ط¯ظ‚ظٹظ‚ط©: $e');
      }
      r.cntNotes.clear();
      r.cntOk = true;
    }
    final col = MCollector();
    _tmStrips = 0;
    _tmOcr = 0;
    var nPages = 0, tRaster = 0;
    final swAll = Stopwatch()..start();
    double? dpiCache; // ظ…ظ‚ط§ط³ ط§ظ„طµظپط­ط§طھ ظپظٹ ط§ظ„ظ…ظ„ظپ ط§ظ„ظ…ظ…ط³ظˆط­ ظˆط§ط­ط¯: ظ†ظ‚ظٹط³ ظ…ط±ط© ط¨ط³
    // طµظپط­ط© طµظپط­ط© (ظ…ط´ ط§ظ„ظ…ظ„ظپ ظƒظ„ظ‡ ظ…ط±ط© ظˆط§ط­ط¯ط©) ط¹ظ„ط´ط§ظ† ط§ظ„ط°ط§ظƒط±ط© ظ…ط§ طھط®ظ„طµط´ ظˆط§ظ„طھط·ط¨ظٹظ‚ ظ…ط§ ظٹظ‚ظپظ„ط´
    for (var pg = 1; pg <= 400; pg++) {
      var got = false;
      try {
        final t0 = swAll.elapsedMilliseconds;
        final dpi = dpiCache ??= await _safeDpi(bytes, pg - 1, 220);
        await for (final page in Printing.raster(bytes, pages: [pg - 1], dpi: dpi)) {
          got = true;
          nPages++;
          final png = await page.toPng();
          tRaster += swAll.elapsedMilliseconds - t0;
          final tp = File('$tmpDir/m_$pg.png');
          await tp.writeAsBytes(png);
          await ocrPageStrips(rec, tp.path, '$tmpDir/m_${pg}_s', col);
          try {
            tp.deleteSync();
          } catch (_) {}
        }
      } catch (e) {
        if (!_pageEnd(e) || pg == 1) rethrow;
      }
      if (!got) break;
      await Future.delayed(const Duration(milliseconds: 40));
    }
    _dbg('PAGES ${r.pdfName}: $nPages طµظپط­ط©طŒ ط§ظ„ظ…ط¬ظ…ظˆط¹ ${swAll.elapsedMilliseconds ~/ 1000}ط« '
        '(raster+png ${tRaster ~/ 1000}ط«طŒ ط´ط±ط§ط¦ط­ ${_tmStrips ~/ 1000}ط«طŒ OCR ${_tmOcr ~/ 1000}ط«)');
    final sc = col.finish();
    r.count = sc.count;
    r.detail = 'RRN ظƒط§ظ…ظ„ ${sc.full} â€¢ ظ†ط§ظ‚طµ ط§ظ„ظ‚ط±ط§ط،ط© ${sc.partial}${sc.rangeText}';
    if (!sc.agree) {
      r.cntOk = false;
      r.cntNotes.add(sc.seqNote);
    }
    if (sc.otherDay.isNotEmpty) {
      r.cntOk = false;
      r.cntNotes.add('ظ…ط´ طھط¨ط¹ ظ†ظپط³ ط§ظ„ظٹظˆظ…: ${sc.otherDay.join(' طŒ ')}');
    }
    if (sc.nonPayment) {
      r.cntNotes.add('ظپظٹظ‡ ط¥ظٹطµط§ظ„ ط¥ظ‚ظپط§ظ„/ظ…ظ„ط®طµ/ط¥ط¹ط§ط¯ط© ط·ط¨ط§ط¹ط© (ط§طھط¬ط§ظ‡ظ„ ظˆظ…ط´ ظ…ط¹ط¯ظˆط¯)');
    }
    if (r.newName == null && sc.tid != null) {
      // ط§ظ„طھط¹ط±ظٹظپ ط§ظ„ط³ط±ظٹط¹ ظپط´ظ„: ظ†ط³ظ…ظ‘ظٹ ظ…ظ† ظ†طھظٹط¬ط© ط§ظ„ط¹ط¯ ط§ظ„ظƒط§ظ…ظ„ط©
      await _applyName(r, src, sc.tid, sc.day, sc.month);
    } else if (sc.tid != null && terminals[sc.tid] != null && r.orderNum != null &&
        terminals[sc.tid]!.orderNum != r.orderNum) {
      r.cntOk = false;
      r.cntNotes.add('ط§ظ„ط¹ط¯ ظ„ظ‚ظ‰ ظ…ط§ظƒظٹظ†ط© ${terminals[sc.tid]!.orderNum} ط¨ط³ ط§ظ„طھط³ظ…ظٹط© ظ…ط§ظƒظٹظ†ط© ${r.orderNum}طŒ ط±ط§ط¬ط¹');
    }
  }

  Future<void> recount(MResult r) async {
    r.count = -1;
    r.tries = 0;
    saveMState(results);
    if (mounted) setState(() => status = 'ط§طھط¹ظ„ظ‘ظ… ظ„ظ„ط¹ط¯. ط§ط¶ط؛ط· \"ط¨ط¯ط، ط§ظ„ط¹ط¯\".');
  }

  List<String> missingMachines() {
    final got = results.where((r) => !r.excluded).map((r) => r.orderNum).whereType<int>().toSet();
    final out = <String>[];
    for (final n in machineNums()) {
      if (!got.contains(n)) out.add('$n ${terminals[numToTid[n]]!.unitName}');
    }
    return out;
  }

  Future<void> openReport(MResult r) async {
    await Navigator.push(
      context,
      MaterialPageRoute(
        builder: (_) => ReceiptReportPage(
          r: r,
          canEdit: !busy,
          onChanged: () {
            saveMState(results);
            if (mounted) setState(() {});
          },
          onOpenPdf: () => openPdf(r.openPath),
          onRecount: () => recount(r),
          onVerdict: (v) => setVerdict(r, v, toggle: false),
        ),
      ),
    );
    if (mounted) setState(() {});
  }

  Future<void> openOverview() async {
    await Navigator.push(
      context,
      MaterialPageRoute(
        builder: (_) => MachinesReportPage(
          results: results,
          missing: missingMachines(),
          openReport: openReport,
        ),
      ),
    );
    if (mounted) setState(() {});
  }

  Future<void> shareLog() async {
    final f = File('$baseDir/id_debug.txt');
    if (!f.existsSync()) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('ظ…ظپظٹط´ ظ„ظˆط¬ ظ„ط³ظ‡. ط´ط؛ظ‘ظ„ ط§ظ„ظپط­طµ ط§ظ„ط£ظˆظ„.')),
      );
      return;
    }
    await Share.shareXFiles([XFile(f.path, mimeType: 'text/plain')], text: 'ظ„ظˆط¬ ط¥ظٹطµط§ظ„ط§طھ ط§ظ„ط¯ظپط¹');
  }

  @override
  Widget build(BuildContext context) {
    return PopScope(
      canPop: !busy,
      onPopInvokedWithResult: (didPop, _) async {
        if (!didPop && busy) {
          // ط§ظ„ظپط­طµ ط´ط؛ط§ظ„: ظ…ط§ ظ†ظ‚ظپظ„ط´ ط§ظ„ط´ط§ط´ط©طŒ ظ†ط±ط¬ظ‘ط¹ ط§ظ„طھط·ط¨ظٹظ‚ ظ„ظ„ط®ظ„ظپظٹط© ط¨ط³ ط¹ظ„ط´ط§ظ† ط§ظ„ظپط­طµ ظٹظƒظ…ظ‘ظ„
          _dbg('BACK while busy -> moveToBack');
          try {
            await _chan.invokeMethod('moveToBack');
          } catch (_) {}
        }
      },
      child: _buildScaffold(context),
    );
  }

  Widget _buildScaffold(BuildContext context) {
    AppSettings.of(context);
    final missing = results.isEmpty ? <String>[] : missingMachines();
    final reviewN = results.where((r) => !r.ok).length;
    final pendCount = results.where((r) => r.pending).length;
    final named = results.where((r) => r.newName != null && !r.excluded).length;
    _syncBadge(machBadge, reviewN);
    machBusyG = busy;
    return Scaffold(
      appBar: AppBar(
        toolbarHeight: 54,
        titleSpacing: 12,
        title: const AppTitle(),
        actions: [
          const HeaderButtons(),
          PopupMenuButton<String>(
            tooltip: tr('ط§ظ„ظ…ط²ظٹط¯'),
            onSelected: (v) {
              if (v == 'appearance') {
                showAppearanceSheet(context);
              } else if (v == 'log') {
                shareLog();
              } else if (v == 'folder') {
                openDest();
              } else if (v == 'src') {
                savePdfDir(null);
                setState(() => pdfDir = null);
              } else if (v == 'out') {
                useDefaultOut();
              } else if (v == 'fast') {
                toggleFastCount();
                setState(() {});
              } else if (v == 'settings') {
                ensureSettingsFile().then((ok) {
                  if (!mounted) return;
                  if (ok) {
                    openFile(settingsPath, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                  } else {
                    ScaffoldMessenger.of(context).showSnackBar(
                      const SnackBar(content: Text('ظ…ظ‚ط¯ط±طھط´ ط£ط¹ظ…ظ„ ظ…ظ„ظپ ط§ظ„ط¥ط¹ط¯ط§ط¯ط§طھ. ط§طھط£ظƒط¯ ظ…ظ† ط¥ط°ظ† ط§ظ„ظˆطµظˆظ„ ظ„ظƒظ„ ط§ظ„ظ…ظ„ظپط§طھ.')),
                    );
                  }
                });
              } else if (v == 'wipe') {
                wipeEverything(context);
              } else if (v == 'clear') {
                clearAll();
              }
            },
            itemBuilder: (_) => [
              if (MediaQuery.sizeOf(context).width < 400)
                PopupMenuItem(
                    value: 'appearance',
                    child: w.Text(AppSettings.of(context).isAr ? 'ط§ظ„ظ…ط¸ظ‡ط± ظˆط§ظ„ظ„ط؛ط©' : 'Theme & language')),
              const PopupMenuItem(value: 'folder', child: Text('ظپطھط­ ظپظˆظ„ط¯ط± ط§ظ„ط¥ط®ط±ط§ط¬')),
              if (pdfDir != null)
                PopupMenuItem(value: 'src', enabled: !busy, child: const Text('ط§ظ„ظ…طµط¯ط±: ط±ط¬ظˆط¹ ظ„ظˆط§طھط³ط§ط¨')),
              if (destDir != defaultMDest)
                PopupMenuItem(value: 'out', enabled: !busy, child: const Text('ط§ظ„ط¥ط®ط±ط§ط¬: ط§ظ„ط§ظپطھط±ط§ط¶ظٹ')),
              PopupMenuItem(
                  value: 'log',
                  child: w.Text(AppSettings.of(context).isAr ? 'ظ…ط´ط§ط±ظƒط© ط§ظ„ظ„ظˆط¬' : 'Share log')),
              const PopupMenuItem(enabled: false, child: Text('ط§ظ„ط¥طµط¯ط§ط± $kAppVersion')),
              PopupMenuItem(value: 'fast', child: Text(fastCount ? 'ط§ظ„ط¹ط¯: ط³ط±ظٹط¹ â†گ ط§ظ„ط¯ظ‚ظٹظ‚ (ط§ظ„ظ‚ط¯ظٹظ…)' : 'ط§ظ„ط¹ط¯: ط¯ظ‚ظٹظ‚ â†گ ط§ظ„ط³ط±ظٹط¹')),
              PopupMenuItem(value: 'settings', child: const Text('ط¥ط¹ط¯ط§ط¯ط§طھ ط§ظ„ظ…ط§ظƒظٹظ†ط§طھ (ط¥ظƒط³ظٹظ„)')),
              PopupMenuItem(value: 'clear', enabled: !busy, child: const Text('ظ…ط³ط­ ط§ظ„ظƒظ„')),
              PopupMenuItem(
                  value: 'wipe',
                  enabled: !busy,
                  child: w.Text(AppSettings.of(context).isAr ? 'ًں§¹ ظ…ط³ط­ ظƒظ„ ط­ط§ط¬ط© (ظƒط§ط´ + ظ…ط¤ظ‚طھ)' : 'ًں§¹ Wipe everything (cache + temp)')),
            ],
          ),
        ],
      ),
      bottomNavigationBar: _bottomBar(
        busy
            ? OutlinedButton.icon(
                onPressed: stopReq ? null : () => setState(() => stopReq = true),
                icon: const Icon(Icons.stop),
                label: const Text('ط¥ظٹظ‚ط§ظپ'),
              )
            : Row(
                children: [
                  Expanded(
                    child: OutlinedButton.icon(
                      onPressed: () => start(),
                      icon: const Icon(Icons.drive_file_rename_outline),
                      label: const Text('ظپط­طµ ظˆطھط³ظ…ظٹط©'),
                    ),
                  ),
                  const SizedBox(width: 8),
                  Expanded(
                    child: FilledButton.icon(
                      onPressed: pendCount > 0 ? () => start(resumeOnly: true, count: true) : null,
                      icon: const Icon(Icons.play_arrow),
                      label: Text(pendCount > 0 ? 'ط¨ط¯ط، ط§ظ„ط¹ط¯ ($pendCount)' : 'ط¨ط¯ط، ط§ظ„ط¹ط¯'),
                    ),
                  ),
                ],
              ),
        busy: busy,
        progress: progressValue,
      ),
      body: ListView(
        padding: const EdgeInsets.fromLTRB(8, 6, 8, 8),
        children: [
          _settingsCard([
            _periodRow(days, busy, (d) {
              setState(() => days = d);
              sharedDays.value = d;
            }),
            const Divider(height: 10),
            _pathRow(Icons.folder_open, 'ط§ظ„ظ…طµط¯ط±', pdfDir ?? 'ظˆط§طھط³ط§ط¨ - ظ…ط³طھظ†ط¯ط§طھ (طھظ„ظ‚ط§ط¦ظٹ)',
                busy ? null : pickDir,
                onReset: pdfDir != null && !busy
                    ? () {
                        savePdfDir(null);
                        setState(() => pdfDir = null);
                      }
                    : null),
            _pathRow(Icons.add_photo_alternate, 'ط§ط®طھظٹط§ط± PDF ظٹط¯ظˆظٹط§ظ‹', 'ظپطھط­ ظ…ط¹ط±ط¶ ط§ظ„ظ…ظ„ظپط§طھ', busy ? null : pickManualPdfs),
            _pathRow(Icons.drive_file_move_outline, 'ط§ظ„ط¥ط®ط±ط§ط¬',
                '$destDir${destDir == defaultMDest ? ' (ط§ظپطھط±ط§ط¶ظٹ)' : ''}', busy ? null : pickOut,
                onReset: destDir != defaultMDest && !busy ? useDefaultOut : null),
          ]),
          _statusBox(status),
          const SizedBox(height: 6),
          _filterRow(reviewOnly, results.length, reviewN, (v) => setState(() => reviewOnly = v)),
          const SizedBox(height: 6),
          Wrap(
            spacing: 6,
            runSpacing: 6,
            children: [
              _mini(Icons.chat, 'ط¥ط±ط³ط§ظ„ ط§ظ„ظƒظ„ ($named)',
                  (busy && !namingDone) || named == 0 ? null : shareAll,
                  filled: true),
              _mini(Icons.insights, _l(context, 'ط§ظ„طھظ‚ط±ظٹط±', 'Report'), results.isEmpty ? null : openOverview,
                  filled: true),
              if (reviewOnly && reviewN > 0)
                _mini(Icons.done_all, 'ط§ظ„ط¨ط§ظ‚ظٹ طµط­ ($reviewN)', busy ? null : markRestOk),
            ],
          ),
          Padding(
            padding: const EdgeInsets.only(top: 4, bottom: 2),
            child: Text('ط§ط³ط­ط¨ ط£ظٹ ظ…ظ„ظپ ظ„طھط¹ظ„ظٹظ…ظ‡: طµط­ âœ“ ط£ظˆ ط؛ظ„ط· âœ—',
                style: TextStyle(fontSize: 11, color: Theme.of(context).colorScheme.onSurfaceVariant)),
          ),
          if (missing.isNotEmpty) ...[
            const SizedBox(height: 10),
            Container(
              padding: const EdgeInsets.all(10),
              decoration: BoxDecoration(
                color: _tintBg(context, Colors.amber),
                border: Border.all(color: _tintBorder(context, Colors.amber)),
                borderRadius: BorderRadius.circular(8),
              ),
              child: Text('âڑ ï¸ڈ ظ…ط§ظƒظٹظ†ط§طھ ظ…ط§ ظˆطµظ„ط´ ظ…ظ„ظپظ‡ط§ (${missing.length}): ${missing.join(' طŒ ')}',
                  style: const TextStyle(fontSize: 13)),
            ),
          ],
          const SizedBox(height: 10),
          for (final r in results.where((r) => !reviewOnly || !r.ok))
            Builder(builder: (context) {
              final match = _batchMatch(r);
              return Dismissible(
                key: ValueKey('t_${r.key}'),
                confirmDismiss: (dir) async {
                  setVerdict(r, dir == DismissDirection.startToEnd ? 1 : -1, toggle: false);
                  return false;
                },
                background: Container(
                  color: _tintBg(context, Colors.green, strong: true),
                  alignment: AlignmentDirectional.centerStart,
                  padding: const EdgeInsets.symmetric(horizontal: 20),
                  child: const Text('طµط­ âœ“', style: TextStyle(fontWeight: FontWeight.bold)),
                ),
                secondaryBackground: Container(
                  color: _tintBg(context, Colors.red, strong: true),
                  alignment: AlignmentDirectional.centerEnd,
                  padding: const EdgeInsets.symmetric(horizontal: 20),
                  child: const Text('ط؛ظ„ط· âœ—', style: TextStyle(fontWeight: FontWeight.bold)),
                ),
                child: Card(
                  margin: const EdgeInsets.symmetric(vertical: 2),
                  color: r.excluded
                      ? _tintBg(context, Colors.grey, strong: true)
                      : (match == null
                          ? null
                          : (match ? _tintBg(context, Colors.green) : _tintBg(context, Colors.red))),
                  child: ListTile(
                    contentPadding: const EdgeInsets.symmetric(horizontal: 12),
                    onTap: () => openReport(r),
                    title: Text(r.newName ?? 'طں ${r.pdfName}',
                        style: TextStyle(
                            fontWeight: FontWeight.bold,
                            decoration: r.excluded ? TextDecoration.lineThrough : null)),
                    subtitle: w.Text(
                        '${r.verdict > 0 ? _l(context, 'âœ“ ط§طھط±ط§ط¬ط¹: طµط­\n', 'âœ“ Reviewed: OK\n') : ''}'
                        '${r.pending && busy ? _l(context, 'ط¨ظٹطھط¹ط¯ظ‘...', 'Counting...') : _mSummary(context, r)}'
                        '${r.notes.isEmpty ? '' : '\n${_l(context, 'âڑ ï¸ڈ ${r.notes.length} ظ…ظ„ط§ط­ط¸ط© - ط§ط¶ط؛ط· ظ„ظپطھط­ ط§ظ„طھظ‚ط±ظٹط±', 'âڑ ï¸ڈ ${r.notes.length} note(s) - tap to open the report')}'}'),
                    trailing: SizedBox(
                      width: 118,
                      child: Row(
                        mainAxisSize: MainAxisSize.min,
                        children: [
                          SizedBox(
                            width: 66,
                            child: TextField(
                              controller: r.ctl,
                              keyboardType: TextInputType.number,
                              textAlign: TextAlign.center,
                              decoration: InputDecoration(
                                isDense: true,
                                contentPadding:
                                    const EdgeInsets.symmetric(horizontal: 4, vertical: 10),
                                hintText: _batchAutoOf(r)?.toString() ?? tr('ط§ظ„طµط­'),
                                hintStyle: const TextStyle(fontSize: 13),
                                border: const OutlineInputBorder(),
                                filled: match != null,
                                fillColor: match == null
                                    ? null
                                    : (match ? _tintBg(context, Colors.green, strong: true) : _tintBg(context, Colors.red, strong: true)),
                              ),
                              onChanged: (v) {
                                saveMState(results);
                                setState(() {});
                              },
                            ),
                          ),
                          PopupMenuButton<String>(
                            tooltip: tr('ط§ظ„ظ…ط²ظٹط¯'),
                            padding: EdgeInsets.zero,
                            onSelected: (v) {
                              if (v == 'report') {
                                openReport(r);
                              } else if (v == 'open') {
                                openPdf(r.openPath);
                              } else if (v == 'share') {
                                shareOne(r);
                              } else if (v == 'recount') {
                                recount(r);
                              } else if (v == 'assign') {
                                assignMachine(r);
                              } else if (v == 'clear') {
                                setVerdict(r, 0, toggle: false);
                              }
                            },
                            itemBuilder: (_) => [
                              PopupMenuItem(
                                  value: 'report',
                                  child: w.Text(_l(context, 'طھظ‚ط±ظٹط± ط§ظ„ط¥ظٹطµط§ظ„', 'Receipt report'))),
                              PopupMenuItem(
                                  value: 'open',
                                  child: w.Text(_l(context, 'ظپطھط­ ط§ظ„ظ…ظ„ظپ', 'Open file'))),
                              PopupMenuItem(
                                  value: 'share',
                                  enabled: r.newName != null && !r.excluded,
                                  child: const Text('ط¥ط±ط³ط§ظ„ ط§ظ„ظ…ظ„ظپ')),
                              PopupMenuItem(
                                  value: 'assign', enabled: !busy, child: const Text('طھط­ط¯ظٹط¯ ط§ظ„ظ…ط§ظƒظٹظ†ط©')),
                              PopupMenuItem(
                                  value: 'recount', enabled: !busy, child: const Text('ط¥ط¹ط§ط¯ط© ط§ظ„ط¹ط¯')),
                              if (r.verdict != 0)
                                const PopupMenuItem(value: 'clear', child: Text('ط¥ظ„ط؛ط§ط، ط¹ظ„ط§ظ…ط© طµط­/ط؛ظ„ط·')),
                            ],
                          ),
                        ],
                      ),
                    ),
                  ),
                ),
              );
            }),
        ],
      ),
    );
  }
}
