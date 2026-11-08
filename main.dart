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

// بيختار DPI بحيث الضلع الأطول ما يزيدش عن ~2500 بكسل (PDF الممسوح بيبقى مقاسه كبير وبيخلّي الذاكرة تخلص)
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

// تاريخ إيصال في المستقبل = قراءة غلط
bool _plausibleDay(String? d, String? m) {
  if (d == null || m == null) return true;
  final di = int.tryParse(d), mi = int.tryParse(m);
  if (di == null || mi == null) return false;
  final n = DateTime.now();
  final dt = DateTime(n.year, mi, di);
  return !dt.isAfter(DateTime(n.year, n.month, n.day).add(const Duration(days: 1)));
}

bool _pageEnd(Object e) => e.toString().contains('Invalid page index');

// سجل تشخيصي بيتكتب في id_debug.txt (بيساعد نعرف ليه ملف اتخطّى أو التطبيق خرج)
void _dbg(String m) {
  try {
    final t = DateTime.now().toIso8601String().substring(11, 19);
    File('$baseDir/id_debug.txt').writeAsStringSync('[$t] $m\n', mode: FileMode.append);
  } catch (_) {}
}

// ملف النتيجة اللي التطبيق نفسه بيطلّعه (نتيجة 5.11.pdf) بس، مش أي اسم تاني
bool _isAppReportName(String n) => RegExp(
        r'^(نتيجه|ناتج|الاسكندريه)\s*(\d{1,2}[.\-_]\d{1,2}|اليوم)(\s*\(\d+\))?\.pdf$')
    .hasMatch(_normAr(n));

// أي ملف إخراج طلّعه التطبيق (إكسيل أو PDF): "نتيجة 5.11.xlsx" (والاسم القديم لسه متعرّف عليه)
bool _isOutputFileName(String n) => RegExp(
        r'^(نتيجه|ناتج|الاسكندريه)\s*(\d{1,2}[.\-_]\d{1,2}|اليوم)(\s*\(\d+\))?\.(pdf|xlsx)$')
    .hasMatch(_normAr(n));

// ───────── عناصر واجهة مدمجة (علشان الشاشة ما تبقاش زحمة) ─────────
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
    onLongPress: onReset, // ضغطة مطوّلة = رجوع للافتراضي
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

// رقم الإصدار: بييجي من GitHub Actions وقت البناء (APP_VERSION + رقم البناء)
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

// بيقرا ملف الإعدادات (الماكينات / كلمات الوحدات / أكواد التاجر / تصحيح الأرقام). لو مفيش ملف الإعدادات فاضية.
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
    // ترمنلات وهمية اتسجلت غلط من الـ MERCHANT ID (رقمها جزء من كود تاجر): بنشيلها
    terminals.removeWhere((tid, _) =>
        tid.length == 8 && merchantCodeMap.keys.any((k) => k.length >= 10 && k.contains(tid)));
    rebuildTerminalIndexes();
  } catch (e) {
    _dbg('Excel load error: $e');
    _loadSettingsFromJson();
  }
}

// نسخة احتياطية: لو قراءة ملف الإكسيل فشلت بنقرا آخر إعدادات اتحفظت (JSON)
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

// بيكتب الماكينات والوحدات اللي اتعرفت (من الإيصالات أو من الملف) في ملف الإعدادات
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

// قالب شيت النتيجة الفاضي (عناوين بس) لو مفيش قالب مضاف للتطبيق ولا في الفولدر
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

// بيتنده أول ما التطبيق يفتح (ولما يرجع من شاشة الإذن): بينزّل شيت الإعدادات وشيت النتيجة تلقائي ويحمّل الإعدادات
Future<void> prepareDefaultFiles() async {
  try {
    if (await Permission.manageExternalStorage.isGranted) {
      await ensureDefaultFiles();
      await ensureSettingsFile();
    }
  } catch (_) {}
  await loadSettingsFromExcel();
}

// أول مرة: بينسخ ملفي الإعدادات والقالب (المرفوعين مع المشروع) لمجلد التطبيق لو مش موجودين
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

// بيعمل ملف الإعدادات فاضي (بالعناوين بس) لو مش موجود، عشان المستخدم يملاه
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
      'Terminals: Order = رقم الماكينة (1,2,3...) | TID = رقم الترمنل المكتوب على الإيصال | Row = رقم صف الوحدة في ملف الإكسيل | Unit_Name = اسم الوحدة',
      'Name_Hints: كلمات (بالإنجليزي، مفصولة بفاصلة) لو ظهرت في الإيصال تتعرف الوحدة',
      'Merchant_Codes: كود التاجر -> رقم الصف',
      'TID_Fixes: رقم ترمنل بتقرأه OCR غلط -> الرقم الصح',
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
const String kAppAuthor = 'Ahmed Ramadan - Alexandria';

// ───────────────────── الإعدادات: اللغة + الدارك مود ─────────────────────
class AppSettings extends ChangeNotifier with WidgetsBindingObserver {
  ThemeMode themeMode = ThemeMode.system; // system = حسب النظام
  String? langCode; // null = حسب لغة النظام
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
    if (langCode == null) notifyListeners(); // بنتبع لغة النظام
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

  // أي ويدجت بتنده عليها بتتسجّل وبتتحدّث لوحدها لما اللغة تتغير
  static AppSettings of(BuildContext context) {
    final s = context.dependOnInheritedWidgetOfExactType<_LangScope>();
    return s?.notifier ?? appSettings;
  }
}

final AppSettings appSettings = AppSettings();

// لون هوية التطبيق (A2R): أزرق ملكي. بيتضاف في آخر القايمة علشان اللي اختار لون قبل كده يفضل لونه.
const Color kBrandSeed = Color(0xFF2B59C3);
const int kDefaultSeed = 6; // مكان kBrandSeed في kSeeds
const List<Color> kSeeds = [Colors.teal, Colors.blue, Colors.indigo, Colors.green, Colors.orange, Colors.purple, kBrandSeed];
const List<String> kSeedNamesAr = ['تيل', 'أزرق', 'نيلي', 'أخضر', 'برتقالي', 'بنفسجي', 'A2R (الهوية)'];
const List<String> kSeedNamesEn = ['Teal', 'Blue', 'Indigo', 'Green', 'Orange', 'Purple', 'A2R (brand)'];

// ───────── حالة مشتركة بين التابين ─────────
final ValueNotifier<Map<int, int>> closingTxn = ValueNotifier<Map<int, int>>(<int, int>{}); // رقم الماكينة -> عدد معاملات الباتش
final ValueNotifier<int> wipeSignal = ValueNotifier<int>(0); // بيتغير لما "مسح كل حاجة" يخلص
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

// تأكيد قبل الإرسال لو فيه حاجات لسه مفتوحة
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

// مسح كل حاجة: ملفات الإخراج + الصور + نتايج الفحص + التصحيحات + اللوج + الكاش والملفات المؤقتة (الإعدادات واللغة والثيم بيفضلوا)
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
          ? 'هيتمسح نهائي:\n• ملفات الإخراج (الإكسيل والـ PDF)\n• الصور المرتبة والنسخ المتسمية لإيصالات الدفع\n• نتايج الفحص والتصحيحات المحفوظة\n• اللوج\n• الكاش والملفات المؤقتة\n\nمش هيتمس: ملفات واتساب الأصلية، والإعدادات (اللغة والثيم والفولدرات).'
          : 'This will permanently delete:\n• Output files (Excel and PDF)\n• Sorted images and renamed payment receipt copies\n• Scan results and saved corrections\n• The log\n• Cache and temporary files\n\nNot touched: original WhatsApp files and settings (language, theme, folders).'),
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
  // نسخ إيصالات الدفع المتسمية لو الفولدر مخصص
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
  // الكاش والملفات المؤقتة بتاعة التطبيق
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
      content: w.Text(ar ? 'تم المسح: $files ملف ($mb ميجا) 🧹' : 'Wiped: $files files ($mb MB) 🧹')));
}

final ValueNotifier<int> sharedDays = ValueNotifier<int>(0); // الفترة (اليوم/يومين/...) واحدة للتابين
final ValueNotifier<int> homeBadge = ValueNotifier<int>(0); // عدد اللي محتاج مراجعة في التقفيل
final ValueNotifier<int> machBadge = ValueNotifier<int>(0); // عدد اللي محتاج مراجعة في إيصالات الدفع

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

// ───────────────────── الترجمة (عربي ← إنجليزي) ─────────────────────
// النصوص مكتوبة بالعربي في الكود، وبتتترجم وقت العرض لما اللغة تكون إنجليزي.
// أي نص عربي مش موجود في القاموس بيفضل زي ما هو.
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
  return out.replaceAll('،', ',').replaceAll('؟', '?');
}

const Map<String, String> _kEnExact = {
  'اليوم': 'Today',
  'الكل': 'All',
  'تمام': 'OK',
  'لا': 'No',
  'إلغاء': 'Cancel',
  'مسح': 'Delete',
  'حفظ': 'Save',
  'عرض': 'View',
  'الصح': 'Correct',
  'تغيير': 'Change',
  'تجريبي': 'Dev',
  'يومين': '2 days',
  '؟': '?',
  'صح ✓': 'Correct ✓',
  'غلط ✗': 'Wrong ✗',
  'المزيد': 'More',
  'المصدر': 'Source',
  'الإخراج': 'Output',
  'واتساب': 'WhatsApp',
  'إكسيل': 'Excel',
  'تدوير': 'Rotate',
  'السابق': 'Previous',
  'التالي': 'Next',
  'إيقاف': 'Stop',
  'مراجعة': 'Review',
  'ماكينة': 'Machine',
  'نتيجة': 'Result'
};

const List<List<String>> _kEnRules = [
  [r'(\d+) من (\d+)', '{1} of {2}'],
  [r'إيصال (\d+) بتاريخ', 'Receipt {1} dated'],
  [r'ناقص (\d+) ماكينة', 'missing {1} machine(s)'],
  [r'مراجعة (\d+) / (\d+)', 'Review {1} / {2}'],
  [r'فاضل (\d+) لسه', '{1} remaining']
];

// مرتبة من الأطول للأقصر علشان العبارات الطويلة تتبدّل الأول
const List<List<String>> _kEnPhrases = [
  ['Terminal ID مش واضح، الوحدة اتحددت من Merchant ID/الاسم جوه الـ PDF، راجع', 'Terminal ID unclear; unit determined from Merchant ID/name inside the PDF, review'],
  ['الملف ده مش واضح إنه إيصال ماكينة (قراءة ضعيفة للـ Terminal ID)، راجعه', 'This file does not look like a machine receipt (weak Terminal ID reading), review it'],
  ['الوحدة اتحددت بالاسم/الميرشانت فقط - اتأكد من الماكينة واضغط حفظ', 'Unit was determined by name/merchant only - verify the machine and tap Save'],
  ['اضغط "بدء المعالجة الشاملة" لسحب الصور وتوليد الإكسيل والـ PDF.', 'Tap "Run full processing" to collect the images and generate the Excel and PDF.'],
  ['التطبيق اتجاهل الصورة دي لأن شكلها إيصال دفع أو مستند مش باتش', 'The app ignored this image because it looks like a payment receipt or a document, not a batch'],
  ['اضغط "فحص وإعادة تسمية" لفحص إيصالات الدفع من مستندات واتساب.', 'Tap "Scan & rename" to scan payment receipts from WhatsApp documents.'],
  ['مفيش صور (في الفترة المحددة) ولا ملفات PDF في المجلد المختار', 'No images (in the selected period) and no PDF files in the selected folder'],
  ['العد وقف أكتر من مرة (التطبيق اتقفل). من ⋮ اختار إعادة العد', 'Counting stopped more than once (the app closed). From ⋮ choose Recount'],
  ['مش محسوب في الإجمالي لحد ما تقرر. زوّم على الباتش وقارن', 'Not counted in the total until you decide. Zoom in on the batch and compare'],
  ['اتعمل لها PDF منفصل لكل يوم، ومش داخلة في إجماليات يوم', 'a separate PDF was created for each day, and they are not included in the totals of day'],
  ['زوّم وشوف: لو باتش حقيقي ادخله وكمّل المبلغ والماكينة', 'Zoom in and check: if it is a real batch, enter it and fill in the amount and machine'],
  ['الـ PDF اتعرّف على نفس ماكينة ملف تاني، اتسمّى ماكينة', 'The PDF was identified as the same machine as another file, named machine'],
  ['وهتتمسح نتيجة الفحص. (ملفات واتساب الأصلية مش هتتمس)', 'The scan result will also be deleted. (Original WhatsApp files will not be touched)'],
  ['ظهرت في صورة تانية بمبلغ مختلف - اتحسبت كباتش إضافي', 'appeared in another image with a different amount - counted as an extra batch'],
  ['فيه إيصال إقفال/ملخص/إعادة طباعة (اتجاهل ومش معدود)', 'Contains a closing/summary/reprint receipt (ignored, not counted)'],
  ['فولدر الإخراج ده عام، فهيتمسح بس ملفات التطبيق منه', 'this output folder is shared, so only the app files will be deleted from it'],
  ['عدد إيصالات الدفع: لسه ما اتعدّش (اضغط بدء العد)', 'Payment receipts count: not counted yet (tap Start counting)'],
  ['مش لاقي فولدر مستندات واتساب. اختار فولدر بإيدك', 'Could not find the WhatsApp documents folder. Pick a folder manually'],
  ['اتجاهلته: شكله إيصال دفع/مستند مش باتش - راجعه', 'Ignored: looks like a payment receipt/document, not a batch - review it'],
  ['ملف لسه ما اتعدّش، اضغط "فحص" وهيكمّل من مكانه', 'file(s) not counted yet, tap "Scan" and it will continue where it left off'],
  ['يجب منح إذن الوصول إلى جميع الملفات للمتابعة', 'You must grant all-files access to continue'],
  ['لا توجد صور في مجلد الواتساب بالفترة المحددة', 'No images in the WhatsApp folder for the selected period'],
  ['تعذر العثور على مجلد صور الواتساب تلقائياً', 'Could not find the WhatsApp images folder automatically'],
  ['مش قادر أقرا المبلغ أو الوحدة - اضغط تعديل', 'Cannot read the amount or the unit - tap Edit'],
  ['ممكن إيصالات ناقصة أو مش اتقرت، راجع الملف', 'possibly missing receipts or unread ones, check the file'],
  ['ملف لسه ما اتعدّش. اضغط "بدء العد" لما تحب', 'file(s) not counted yet. Tap "Start counting" when you want'],
  ['الماكينة/الوحدة اتحددت تلقائياً - اتأكد', 'Machine/unit was determined automatically - verify'],
  ['مشكوك إنه مكرر - مش محسوب لحد ما تراجعه', 'Possible duplicate - not counted until reviewed'],
  ['صورة، ولم يتم العثور على إيصالات مطابقة', 'image(s), and no matching receipts were found'],
  ['صفحة فشلت قراءتها، العدد ممكن يكون ناقص', 'page(s) failed to read, the count may be incomplete'],
  ['كل الملفات في الفترة دي اتفحصت قبل كده', 'All files in this period were already scanned'],
  ['جاري توليد الملفات وتحديث الشيتات...', 'Generating files and updating sheets...'],
  ['لازم تسمح بالوصول لكل الملفات الأول', 'You need to allow access to all files first'],
  ['الملف مش موجود. شغّل المعالجة الأول', 'File not found. Run the processing first'],
  ['مراجعة التكرار وترقيم الماكينات...', 'Checking duplicates and numbering machines...'],
  ['الوحدة اتحددت بالاسم/الميرشانت فقط', 'Unit was determined by name/merchant only'],
  ['اسحب أي ملف لتعليمه: صح ✓ أو غلط ✗', 'Swipe any file to mark it: correct ✓ or wrong ✗'],
  ['مقدرتش أفتح الفولدر. المسار اتنسخ', 'Could not open the folder. Path copied'],
  ['تقرير الإيصالات (الرسمي)', 'Receipts report (official)'],
  ['لم يتم إنشاء ملفات بعد لمشاركتها', 'No files have been created yet to share'],
  ['ولا مضاف للتطبيق)، فمتعملش إكسيل', 'nor added to the app), so no Excel was created'],
  ['مفيش صور. بدوّر على ملف PDF اسمه', 'No images. Looking for a PDF file named'],
  ['مش قادر أعرف الماكينة من الـ PDF', 'Cannot identify the machine from the PDF'],
  ['ملفات المراجعة غير متاحة حالياً', 'Review files are not available right now'],
  ['مشكوك إنه مكرر (نفس مبلغ ماكينة', 'Possible duplicate (same amount as machine'],
  ['مش قادر أعرف التاريخ من الـ PDF', 'Cannot identify the date from the PDF'],
  ['استبعدته قبل كده - تقدر ترجّعه', 'excluded earlier - you can restore it'],
  ['المرشح التاني من قراءة الـ PDF', 'second candidate from reading the PDF'],
  ['مفيش لوج لسه. شغّل الفحص الأول', 'No log yet. Run the scan first'],
  ['جاري البحث عن مجلدات الصور...', 'Searching for image folders...'],
  ['إعادة تسمية الصور وترتيبها...', 'Renaming and sorting images...'],
  ['اضغط "بدء العد" لعدّ إيصالاته', 'Tap "Start counting" to count its receipts'],
  ['صفحات الطريقتين اختلفوا فيها:', 'Pages where the two methods differ:'],
  ['اتعلّم للعد. اضغط "بدء العد".', 'Marked for counting. Tap "Start counting".'],
  ['مشكوك إنه مكرر: نفس الماكينة', 'Possible duplicate: same machine'],
  ['اضغط "بدء العد" لعدّ إيصالات', 'Tap "Start counting" to count receipts of'],
  ['صفحات ما لقيتش فيها إيصالات:', 'Pages with no receipts found:'],
  ['العد: سريع ← الدقيق (القديم)', 'Counting: fast, then precise (old)'],
  ['عدد إيصالات الدفع: بيتعدّ...', 'Payment receipts count: counting...'],
  ['تم توليد ملف الـ PDF المجمع', 'Combined PDF file generated'],
  ['رقم الماكينة اتحدد تلقائياً', 'Machine number was determined automatically'],
  ['مش مكرر - احتسبه باتش إضافي', 'Not a duplicate - count it as an extra batch'],
  ['إيصال (رقمه مش واضح) بتاريخ', 'Receipt (number unclear) dated'],
  ['مفيش ملفات مسمّاة لمشاركتها', 'No named files to share'],
  ['مفيش ملفات PDF في الفترة دي', 'No PDF files in this period'],
  ['مش قادر أكتب في الفولدر ده', 'Cannot write to this folder'],
  ['مفيش إيصالات محتاجة مراجعة', 'No receipts need review'],
  ['صورة متجاهلة - محتاجة قرار', 'Ignored image - needs a decision'],
  ['مفيش برنامج يفتح الملف ده', 'No app can open this file'],
  ['فيه إيصالات بتواريخ تانية', 'There are receipts with other dates'],
  ['فحص الصور المقلوبة (أبطأ)', 'Check rotated images (slower)'],
  ['جاري فحص إيصالات الدفع...', 'Scanning payment receipts...'],
  ['واتساب - مستندات (تلقائي)', 'WhatsApp - documents (auto)'],
  ['تقرير الإيصالات', 'Receipts report'],
  ['جاري تجهيز صفحات الـ PDF', 'Preparing PDF pages'],
  ['الإيصال ده مستبعد (مكرر)', 'This receipt is excluded (duplicate)'],
  ['مفيش برنامج يفتح الـ PDF', 'No app can open the PDF'],
  ['مش قادر أفتح الفولدر ده', 'Cannot open this folder'],
  ['انتهى فحص إيصالات الدفع', 'Payment receipts scan finished'],
  ['جاري البحث عن ملفات PDF', 'Searching for PDF files'],
  ['مش محسوب لحد ما تراجعه', 'not counted until you review it'],
  ['إجمالي بعيد عن المعتاد', 'Total far from usual'],
  ['القراءة التانية للمبلغ', 'The second amount reading'],
  ['صورة من الصور المرتبة', 'image(s) from the sorted images'],
  ['جاري فحص الإيصالات...', 'Scanning receipts...'],
  ['ولا يوجد ملف PDF باسم', 'and there is no PDF file named'],
  ['لو المبلغ صح اضغط حفظ', 'if the amount is correct, tap Save'],
  ['ده إيصال أنهي ماكينة؟', 'Which machine is this receipt for?'],
  ['الملف الأصلي مش موجود', 'Original file not found'],
  ['ماكينات ما وصلش ملفها', 'Machines with no file received'],
  ['تعديل بيانات الإيصال', 'Edit receipt data'],
  ['راجعه من زر المراجعة', 'review it using the review button'],
  ['المصدر: رجوع لواتساب', 'Source: back to WhatsApp'],
  ['بدء المعالجة الشاملة', 'Run full processing'],
  ['بينما أغلب الإيصالات', 'while most receipts are'],
  ['افتح الفولدر بواسطة', 'Open folder with'],
  ['المبلغ يحتاج مراجعة', 'Amount needs review'],
  ['إيصال مكرر للماكينة', 'Duplicate receipt for machine'],
  ['مقدرتش أدوّر الصورة', 'Could not rotate the image'],
  ['اتأكد منه واضغط حفظ', 'verify it and tap Save'],
  ['للمراجعة (مش معدود)', 'For review (not counted)'],
  ['راجعها لو العدد مهم', 'check them if the count matters'],
  ['العد: دقيق ← السريع', 'Counting: precise, then fast'],
  ['متصفح ملفات النظام', 'System file browser'],
  ['حفظ وتحديث الملفات', 'Save and update files'],
  ['بقرأ صفحات ملف PDF', 'Reading pages of PDF file'],
  ['راجع المبلغ يدوياً', 'check the amount manually'],
  ['الإخراج: الافتراضي', 'Output: default'],
  ['اختيار هذا الفولدر', 'Choose this folder'],
  ['مستبعد (علامة غلط)', 'Excluded (marked wrong)'],
  ['عدد إيصالات الدفع:', 'Payment receipts count:'],
  ['إلغاء علامة صح/غلط', 'Clear correct/wrong mark'],
  ['مش موجود في فولدر', 'not found in folder'],
  ['الوحدة غير معروفة', 'Unknown unit'],
  ['فتح فولدر الإخراج', 'Open output folder'],
  ['التاريخ صح - سيبه', 'The date is correct - keep it'],
  ['مش تبع نفس اليوم:', 'Not from the same day:'],
  ['بس التسمية ماكينة', 'but the naming says machine'],
  ['لوج إيصالات الدفع', 'Payment receipts log'],
  ['الأيام اللي فاتت', 'previous days'],
  ['تم توليد الإكسيل', 'Excel generated'],
  ['ملاحظات للمراجعة', 'Notes for review'],
  ['المبلغ غير مقروء', 'Amount unreadable'],
  ['مش باتش - تجاهله', 'Not a batch - ignore it'],
  ['القراءة التانية:', 'Second reading:'],
  ['أرقام مش لاقيها:', 'Numbers not found:'],
  ['ملف على إنهم صح؟', 'file(s) as correct?'],
  ['ملف ونتيجة الفحص', 'file(s) and the scan result'],
  ['اتأكد من المبلغ', 'check the amount'],
  ['مفيش حاجة للمسح', 'Nothing to delete'],
  ['المبلغ الإجمالي', 'Total amount'],
  ['الإجمالي العام:', 'Grand total:'],
  ['المصدر: ملف PDF', 'Source: PDF file'],
  ['واتساب (تلقائي)', 'WhatsApp (auto)'],
  ['وحدة غير معروفة', 'Unknown unit'],
  ['ده باتش - ادخله', 'This is a batch - include it'],
  ['اتعرفت كمان في:', 'also identified in:'],
  ['العد لقى ماكينة', 'Counting found machine'],
  ['الملف مش موجود', 'File not found'],
  ['ملف من الإخراج', 'file(s) from the output'],
  ['تقرير المراجعة', 'Review report'],
  ['فشل ملء القالب', 'Failed to fill the template'],
  ['ماكينات ناقصة:', 'Missing machines:'],
  ['مشكوك إنه مكرر', 'possible duplicate'],
  ['رجّعه (احتسبه)', 'Restore (count it)'],
  ['اكتب مبلغ صحيح', 'Enter a valid amount'],
  ['مكرر - استبعده', 'Duplicate - exclude it'],
  ['اختار الماكينة', 'Select machine'],
  ['اتسمّى، إجمالي', 'named, total'],
  ['تحديد الماكينة', 'Set machine'],
  ['إيصالات الدفع', 'Payment receipts'],
  ['إجمالي ماكينة', 'Total of machine'],
  ['بعيد عن متوسط', 'is far from the average of'],
  ['مستبعد (مكرر)', 'Excluded (duplicate)'],
  ['رقم الماكينة:', 'Machine number:'],
  ['ماكينات ناقصة', 'Missing machines'],
  ['تاريخ الإيصال', 'Receipt date'],
  ['اتسمّى ماكينة', 'Named machine'],
  ['ملف مسمّى من:', 'named file(s) from:'],
  ['اتسمّى واتنسخ', 'named and copied'],
  ['تعريف الملفات', 'Identifying files'],
  ['عدّ الإيصالات', 'Counting receipts'],
  ['قراءة تانية:', 'Second reading:'],
  ['فشلت القراءة', 'read failed'],
  ['نفس الماكينة', 'Same machine'],
  ['مستند متجاهل', 'Ignored document'],
  ['اختيار فولدر', 'Choose folder'],
  ['فشل التسمية:', 'Naming failed:'],
  ['تمت التسمية:', 'Naming done:'],
  ['فشل التعريف:', 'Identification failed:'],
  ['ناقص القراءة', 'partial read'],
  ['لوج التقفيل', 'Closing log'],
  ['تاريخ مختلف', 'Different date'],
  ['انتهى الفحص', 'Scan finished'],
  ['ونفس المبلغ', 'and the same amount'],
  ['راجع الصورة', 'check the image'],
  ['خليه بتاريخ', 'Set it to date'],
  ['حفظ والتالي', 'Save & next'],
  ['اتوقف. فاضل', 'Stopped. Remaining'],
  ['اكتمل العد:', 'Counting complete:'],
  ['إرسال الملف', 'Send file'],
  ['تغيير اللغة', 'Change language'],
  ['الإسكندرية', 'Alexandria'],
  ['ملء الشاشة', 'Full screen'],
  ['بس التسلسل', 'but the sequence'],
  ['آخر نتيجة:', 'Last result:'],
  ['فحص وتسمية', 'Scan & name'],
  ['إرسال الكل', 'Send all'],
  ['اتراجع: صح', 'Reviewed: correct'],
  ['إعادة العد', 'Recount'],
  ['مختلفة عن', 'differs from'],
  ['اتأكد منه', 'verify it'],
  ['غير مقروء', 'unreadable'],
  ['الباقي صح', 'The rest are correct'],
  ['مسح الكل', 'Delete all'],
  ['غير محدد', 'Unspecified'],
  ['للمراجعة', 'for review'],
  ['الماكينة', 'Machine'],
  ['المقروء:', 'Read:'],
  ['استخدمها', 'Use it'],
  ['فشل العد', 'Counting failed'],
  ['عدّ سريع', 'Quick count'],
  ['بدء العد', 'Start counting'],
  ['التقفيل', 'Closing'],
  ['ماكينات', 'Machines'],
  ['والأغلب', 'while most are'],
  ['المبلغ:', 'Amount:'],
  ['الإصدار', 'Version'],
  ['افتراضي', 'default'],
  ['التسلسل', 'Sequence'],
  ['الفترة', 'Period'],
  ['هتتمسح', 'Will delete'],
  ['تم مسح', 'Deleted'],
  ['ماكينة', 'Machine'],
  ['اكتمل!', 'Done!'],
  ['باتشات', 'batches'],
  ['تاريخه', 'Date'],
  ['مستبعد', 'Excluded'],
  ['تم فحص', 'Scanned'],
  ['بتاريخ', 'dated'],
  ['هتعلّم', 'Mark'],
  ['هيتمسح', 'Will delete'],
  ['الأصل:', 'Original:'],
  ['اتسمّى', 'named'],
  ['مراجعة', 'Review'],
  ['ملفات', 'files'],
  ['إيصال', 'receipt(s)'],
  ['صفحات', 'pages'],
  ['العدد', 'Count'],
  ['مقروء', 'read'],
  ['أيام', 'days'],
  ['قالب', 'Template'],
  ['باتش', 'batch'],
  ['ناقص', 'missing'],
  ['صفحة', 'page(s)'],
  ['صورة', 'image(s)'],
  ['فاضل', 'remaining'],
  ['راجع', 'review'],
  ['الكل', 'All'],
  ['كامل', 'full'],
  ['ملف', 'file(s)'],
  ['ج.م', 'EGP'],
  ['فحص', 'Scan'],
  ['فيه', 'There are'],
  ['لسه', 'yet'],
  ['أو', 'or'],
  ['و', 'and']
];

// ───────────────────── ألوان تتأقلم مع الدارك مود ─────────────────────
bool _isDark(BuildContext c) => Theme.of(c).brightness == Brightness.dark;

Color _tintBg(BuildContext c, MaterialColor m, {bool soft = false, bool strong = false}) {
  if (_isDark(c)) return m.shade900.withAlpha(strong ? 210 : 140);
  return strong ? m.shade300 : (soft ? m.shade50 : m.shade100);
}

Color _tintBorder(BuildContext c, MaterialColor m) => _isDark(c) ? m.shade300 : m.shade800;
Color _tintFg(BuildContext c, MaterialColor m) => _isDark(c) ? m.shade200 : m.shade900;

// ───────────────────── الهيدر: اسم التطبيق + زراير اللغة والدارك ─────────────────────
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

// على الشاشات العريضة: أيقونة منفصلة. على الصغيرة: بتتنقل لقايمة "المزيد" (بند "المظهر واللغة").
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

// ───────────────────── جداول الماكينات والوحدات ─────────────────────
class TerminalInfo {
  final int orderNum;
  final int excelRow;
  final String unitName;
  const TerminalInfo(this.orderNum, this.excelRow, this.unitName);
}

// النسخة العامة: مفيش ماكينات ولا وحدات مكتوبة في الكود. كله بيتقرأ من ملف A2R_Settings.xlsx
final Map<String, TerminalInfo> terminals = {};
final Map<int, String> numToTid = {};
final Map<String, String> tidFix = {};
final Map<String, String> tidAlias = {};
final Map<String, int> merchantCodeMap = {};
final List<MapEntry<List<String>, int>> nameHints = [];
final Map<int, String> rowNames = {};

// بيبني الفهارس (رقم الماكينة -> TID ، رقم الصف -> اسم الوحدة) من الماكينات المحمّلة
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

// ───────────────────── خوارزميات الاستخراج (منقولة من نسخة البايثون) ─────────────────────
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

// سطر نص مع مكانه في الصورة
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

// يرتب النص صفوف (الاسم يمين والرقم جنبه في نفس السطر) زي Tesseract بالظبط
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

// صورة فيها أكتر من إيصال جنب بعض -> نقسمها لأعمدة
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

// ───────────── المبالغ ─────────────
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
  // 1) تطابق تام بين مجموع الكروت و GRAND TOTALS
  for (final r in results) {
    final b = r.bsum;
    if (b != null && b > 0 && r.grand.any((g) => (g - b).abs() < 0.005)) {
      return Decision(b, true);
    }
  }
  // 2) فرق رقم واحد
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
  // 3) إجمالي مكرر
  for (final r in results) {
    final c = mostCommon(r.grand);
    if (c != null && c.value >= 2 && r.bsum == null) return Decision(c.key, false);
  }
  // 4) تصويت
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

// إيصال دفع (RRN / رقم كارت مخفي / Approved ...) ولا باتش كلوز / ملخص / ري برنت؟
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

// لو GRAND TOTAL = مجموع إجماليات الكروت (حتى لو أسماء الكروت اتقرت غلط): نرجّعه مؤكد.
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

// ───────────── الماكينة / الوحدة / التاريخ ─────────────
String? _tidLabeledRaw(String t) {
  final m = RegExp(r"TERMINAL\s*(?:ID)?\D{0,15}?(\d[\d\s/]{6,13}\d)",
          caseSensitive: false)
      .firstMatch(t);
  if (m == null) return null;
  final raw = m.group(1)!.replaceAll('/', '7').replaceAll(RegExp(r'\D'), '');
  // رقم الترمنل 8 أرقام. أي رقم أطول (زي MERCHANT ID اللي OCR أحياناً يحطه جنب TERMINAL) بنتجاهله
  if (raw.length < 8 || raw.length > 9 || raw.startsWith('0011')) return null;
  final d = raw.substring(raw.length - 8);
  return tidFix[d] ?? d;
}
String? _tidLabeled(String t) {
  final d = _tidLabeledRaw(t);
  return d != null && terminals.containsKey(d) ? d : null;
}

// ───────────── تعلّم الماكينات والوحدات من الإيصالات ─────────────
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

// اسم الوحدة المكتوب على الإيصال: بعد كلمة MERCHANT NAME لو موجودة، وإلا أول سطر اسم في أعلى الإيصال
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
  // إيصالات NBE: اسم الوحدة بيبقى سطر بالشكل "EC <الاسم> TRF" (حتى لو OCR دمجه مع سطر تاني)
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
  // اسم الوحدة بيبقى تحت سطر "NATIONAL BANK OF EGYPT" على طول (قبل سطر العنوان)
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

// الـ MERCHANT ID المكتوب على الإيصال (الأكتر تكراراً بين القراءات)
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

// لو الـ MERCHANT ID ده متسجل قبل كده (أو قريب جداً منه بفرق رقم واحد من غلط القراءة) بنرجّع صف وحدته
int? _rowForMerchant(String mid) {
  final exact = merchantCodeMap[mid];
  if (exact != null) return exact;
  for (final e in merchantCodeMap.entries) {
    final k = e.key;
    // مفيش مطابقة تقريبية: أكواد التجار بتاعة الوحدات بتختلف في رقم واحد بس (…029084 / …026084 / …022084)
    if (k.length >= 10 && mid.length > k.length && mid.contains(k)) return e.value;
  }
  return null;
}

// ماكينة معروفة بس الـ MERCHANT ID بتاعها لسه مش متسجل: بنسجّله علشان الماكينات الجاية بنفس الكود تدخل نفس الوحدة
void learnMerchant(String tid, List<String> texts) {
  final t = terminals[tid];
  if (t == null) return;
  final mid = extractMerchantId(texts);
  if (mid == null || _rowForMerchant(mid) != null) return;
  merchantCodeMap[mid] = t.excelRow;
  saveSettingsToExcel();
}

// ماكينة جديدة مكتوبة على الإيصال (Terminal ID) ومش في القايمة: بنسجّلها هي ووحدتها وبنحفظها في ملف الإعدادات
String? discoverTerminal(List<String> texts) {
  final counts = <String, int>{};
  final midNow = extractMerchantId(texts);
  bool inMerchant(String d) => midNow != null && midNow.contains(d);
  for (final t in texts) {
    final d = _tidLabeledRaw(t);
    if (d != null && !inMerchant(d)) counts[d] = (counts[d] ?? 0) + 1;
  }
  if (counts.isEmpty) {
    // لو الـ OCR لخبط ترتيب الأسطر: رقم من 8 أرقام بيبدأ بـ 50 واقف لوحده
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
  // ماكينات بترمنل مختلف لكن نفس الـ MERCHANT ID = نفس الوحدة: بتتحسب مع بعض في صف واحد
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

// عدد معاملات الباتش من إيصال الإقفال: "GRAND TOTALS / Sale (010)" (ولو اتقرا مجموع الكروت بيتطابق معاه)
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
  if (g != null && s != null && g != s) return null; // القراءتين متعارضتين: الأفضل تكتبه بإيدك
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
  int? txn; // عدد معاملات الباتش (من إيصال الإقفال)
  Analysis(this.tid, this.how, this.row, this.unitName, this.val, this.ok,
      this.batches, this.date);
}

// تحليل صورة كاملة (ممكن فيها أكتر من إيصال)
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
    if (b != null && seen.contains(key)) continue; // نفس الباتش اتصور مرتين
    seen.add(key);
    batches.add(r.val);
    txns.add(batchTxnOf(r.texts));
    allOk = allOk && r.ok;
  }
  var total = round2(batches.fold(0.0, (a, b) => a + b));
  var ok = allOk && batches.isNotEmpty;

  // شبكة أمان: لو التقسيم مش مؤكد جرّب الصورة كلها
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
  result.txn = (txns.isNotEmpty && txns.every((e) => e != null)) ? txns.fold<int>(0, (a, e) => a + e!) : null;
  return result;
}

double round2(double val) => double.parse(val.toStringAsFixed(2));

// دمج القراءة التانية (الصورة المحسّنة) في القراءة الأولى.
// بترجع قيمة المبلغ التانية لو متعارضة مع الأولى (علشان تتراجع)، وإلا null.
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
  if ((a.val - b.val).abs() < 0.005) return null; // القراءتين متطابقتين
  if (!a.ok && b.ok) {
    // الأولى مش مؤكدة والتانية مؤكدة بمجموع الكروت -> نعتمد التانية
    a.val = b.val;
    a.ok = true;
    a.batches = [...b.batches];
    a.txn = b.txn ?? a.txn;
    return null;
  }
  if (a.ok && !b.ok) return null; // الأولى مؤكدة، نتجاهل التانية
  a.ok = false; // قراءتين مختلفتين: لازم مراجعة
  return b.val;
}

// بيحوّل صفحات ملف PDF لصور PNG (بيستخدم في حالة مفيش صور)
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

// ───────────────────── مسارات المجلدات ─────────────────────
const String root = '/storage/emulated/0';
final String baseDir = '$root/A2R-POS-Public';
final String rawImagesDir = '$baseDir/raw_images';
String outputDir = '$baseDir'; // ممكن المستخدم يغيّره من الشاشة
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

// ───────────────── معالجة الصور (تدوير / تحسين) - بتشتغل في Isolate ─────────────────
// بتصلّح اتجاه الـ EXIF وبتدوّر الصورة لو deg != 0. بترجع true لو كتبت ملف جديد.
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

// نسخة محسّنة للقراءة التانية: رمادي + كونتراست + تكبير/تصغير لمقاس مناسب للـ OCR
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

// ───────────────── سكانر الباتش: قص الإيصال من الصورة + تعديل المنظور + تحسين (زي كام سكانر) ─────────────────
// الفكرة: الورقة البيضا (الباتش) أفتح من الخلفية (ترابيزة/إيد/أرضية). بنفصل الورقة بحدّ Otsu،
// بنطلّع أكبر مكوّن، بنفرد منظوره (4 أركان)، وبعدين بنشيل الظلال ونبيّض الورقة ونغمّق الكتابة.
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

// خط مستقيم q = m*p + k بأقل مربعات مع استبعاد النقط الشاردة. null لو النقط قليلة.
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

// أركان الورقة [tlx,tly,trx,try,blx,bly,brx,bry] بإحداثيات الصورة الصغيرة
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

  // أصغر مستطيل مايل يحيط بالورقة
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
  final lp = vertical ? ay : ax; // على طول الورقة
  final lq = vertical ? ax : ay; // عرض الورقة
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

// بيرجّع أركان كل باتش في الصورة (بإحداثيات الصورة الأصلية). فاضي = مش عارف يلاقي ورقة.
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
    // الصورة كلها ورقة (سكرين/اسكانر جاهز): بس نحسّنها
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
  // لو في أكتر من باتش في الصورة: رتّبهم شمال ليمين (أو فوق لتحت لو فوق بعض)
  out.sort((a, b) {
    final acx = (a[0] + a[2] + a[4] + a[6]) / 4, bcx = (b[0] + b[2] + b[4] + b[6]) / 4;
    final acy = (a[1] + a[3] + a[5] + a[7]) / 4, bcy = (b[1] + b[3] + b[5] + b[7]) / 4;
    final aw = _dist(a[0], a[1], a[2], a[3]), bw = _dist(b[0], b[1], b[2], b[3]);
    if ((acx - bcx).abs() > math.min(aw, bw) * 0.5) return acx.compareTo(bcx);
    return acy.compareTo(bcy);
  });
  return out;
}

// بيفرد الورقة (تعديل منظور) ويحسّنها: يشيل الظلال، يبيّض الخلفية، يغمّق الكتابة
img.Image? _warpAndEnhance(img.Image im, List<double> q) {
  final wq = (_dist(q[0], q[1], q[2], q[3]) + _dist(q[4], q[5], q[6], q[7])) / 2;
  final hq = (_dist(q[0], q[1], q[4], q[5]) + _dist(q[2], q[3], q[6], q[7])) / 2;
  if (wq < 40 || hq < 40) return null;
  final k = math.min(1.0, 2200.0 / math.max(wq, hq));
  final wo = math.max(8, (wq * k).round());
  final ho = math.max(8, (hq * k).round());

  // قص المنطقة المحيطة بالورقة بس (أسرع كتير) وتصغيرها لو كبيرة
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

  // فرد المنظور (استيفاء ثنائي بين الأركان الأربعة)
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

  // تحسين: نقدّر لون الورقة (أعلى قيمة في كل مربع) ونقسم عليه = شيل الظلال وتبييض الخلفية
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
      // منحنى أغمق للكتابة (gamma) علشان الأرقام تفضل تقيلة وواضحة
      final o = math.pow(_cd((n - 30.0) / (240.0 - 30.0), 0.0, 1.0), 1.6).toDouble();
      final c = (o * 255).round();
      res.setPixelRgb(x, y, c, c, c);
    }
  }
  return res;
}

// بتشتغل في Isolate. بترجّع مسارات الصفحات (صفحة لكل باتش). فاضي = استخدم الصورة الأصلية.
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


// ───────────────── ملء القالب مباشرة جوه ملف الـ xlsx (من غير ما نبوّظ التنسيق) ─────────────────
Future<List<int>?> loadTemplateBytes() async {
  try {
    final mine = File('$baseDir/template.xlsx');
    if (mine.existsSync()) return mine.readAsBytesSync();
    final exact = File('$baseDir/نتيجة 00.xlsx');
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

// بيكتب خلية واحدة محافظ على الـ style بتاعها. build(style, oldFormulaElement) -> xml الخلية
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

// بيبني صفحة النتايج: صف العناوين زي ما هو، وصف لكل وحدة بنفس تنسيق أول صف بيانات في القالب
// (عدد الصفوف = عدد الوحدات اللي اتعرفت من الفحص، بيزيد أو بيقل حسب العدد)
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
    if (f.name == 'xl/calcChain.xml') continue; // بيتعمل من جديد، وكان هيبوّظ الملف بعد تغيير الصفوف
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

// ───────────────── قنوات النظام (واتساب مباشر / فتح فولدر) ─────────────────
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

// ملف PDF تقفيل اسمه فيه تاريخ النهارده (بيقبل 5.11 / 05.11 / 5-11 / 5_11)
bool isTodayAlexPdf(String path) {
  final name = _normAr(baseName(path));
  if (!name.endsWith('.pdf')) return false;
  if (!(name.contains('اسكندري') || name.contains('نتيج') || name.contains('تقفيل') || name.contains('اقفال'))) return false;
  final n = DateTime.now();
  final re = RegExp('(?<!\\d)0?${n.day}[._\\- ]0?${n.month}(?!\\d)');
  return re.hasMatch(name);
}

List<File> findTodayAlexPdfs(List<String> dirs) {
  final out = <File>[];
  final seen = <String>{};
  for (final d in dirs) {
    if (d == outputDir) continue; // ما نقراش الملفات اللي التطبيق نفسه طلّعها
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

// ───────────────────── نموذج الإيصال ─────────────────────
class ReceiptItem {
  String path;
  String? tid;
  int? orderNum;
  int? row;
  String? unitName;
  double amount;
  String? date;
  String key; // اسم الملف الأصلي + حجمه (لحفظ التصحيحات اليدوية)
  bool ok; // المبلغ مؤكد؟
  String how; // terminal | fuzzy | name/merchant | auto-number | conflict | extra-batch | manual
  List<double> batches;
  int? suffix;
  String? origTid;
  bool conflict;
  bool skip;
  double? alt; // قراءة تانية مختلفة للمبلغ (للمراجعة)
  bool dupSuspect; // مشكوك إنه مكرر: بيدخل المراجعة ومش بيتحسب لحد ما تقرر
  bool userExcluded; // انت استبعدته في المراجعة
  String? dupDecision; // 'keep' | 'drop' (قرارك بيتحفظ للمرة الجاية)
  int? dupOf; // رقم الماكينة اللي شبهه
  String? unusual; // إجمالي بعيد عن المعتاد (للمراجعة)
  bool ignoredDoc; // صورة اتجاهلها التطبيق (شكلها إيصال دفع/مستند): بتدخل المراجعة
  String? ign; // 'keep' | 'drop' قرارك في الصورة المتجاهلة
  bool dateOdd; // تاريخه مختلف عن أغلب الإيصالات
  String? dateMain; // التاريخ الأغلب
  String? dateFix; // التاريخ اللي انت أكدته
  int? txn; // عدد معاملات الباتش المقروء من إيصال الإقفال

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

  // داخل في الإجماليات والـ PDF؟
  bool get counts => !skip && !dupSuspect && !userExcluded && !ignoredDoc;

  // الماكينة/الوحدة متحددة بطريقة مش مؤكدة (تلقائي / بالاسم فقط)
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

// ───────────────────── التصحيحات اليدوية (بتفضل محفوظة) ─────────────────────
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

// صورة كان التطبيق متجاهلها: ادخلها كباتش (هتكمّل المبلغ والماكينة بإيدك)
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

// احتساب إيصال (مشكوك أو مستبعد) كباتش إضافي للماكينة
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

// إجمالي ماكينة بعيد جداً عن متوسط الأيام اللي فاتت = غالباً قراءة غلط -> مراجعة
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


// ───────── العد السريع: صور الصفحات الأصلية (JPEG) من الـ PDF الممسوح + عدّ بالهيكل ─────────
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

// بيطلّع صور الصفحات (JPEG) جاهزة من PDF ممسوح (CamScanner وغيره): من غير raster ولا PNG
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

// عدّ الإيصالات بالهيكل: كل إيصال له شعار + Terminal ID + RRN + AMOUNT + Approved (مرة واحدة)
// بنجمع العلامات دي ونقسمها أعمدة، وفي كل عمود أكتر علامة متكررة = عدد الإيصالات
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
  // إزالة التكرار (نفس العلامة اتقرت مرتين)
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
  // أعمدة
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
  // لكل عمود: كل علامة بتظهر مرة واحدة في كل إيصال، فأكبر عدد لأي علامة = عدد الإيصالات في العمود
  // (لو علامة اتقرت ناقصة التانية بتعوّض، ومش بيتأثر بميل الإيصال ولا بترتيب القراءة)
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

// ───────── اختيار تطبيق (مدير ملفات / معرض صور / ملفات PDF) مع إمكانية حفظ الافتراضي ─────────
Map<String, String>? _loadDefaultHandler(String key) {
  try {
    final v = _readCfg()[key];
    if (v is Map) return {for (final e in v.entries) e.key.toString(): e.value.toString()};
  } catch (_) {}
  return null;
}

void _saveDefaultHandler(String key, Map<String, String>? h) {
  try {
    final cfg = _readCfg();
    if (h == null) {
      cfg.remove(key);
    } else {
      cfg[key] = h;
    }
    _writeCfg(cfg);
  } catch (_) {}
}

// بيمسح الاختيارات المحفوظة (مدير الملفات / المعرض / ملفات PDF) علشان تختار من جديد
void resetDefaultHandlers(BuildContext context) {
  for (final k in ['defFolder', 'defImage', 'defPdf']) {
    _saveDefaultHandler(k, null);
  }
  ScaffoldMessenger.of(context).showSnackBar(
    const SnackBar(content: w.Text('اتمسحت التطبيقات الافتراضية. هتختار من جديد المرة الجاية.')),
  );
}

List<Map<String, String>> _toHandlers(dynamic r) {
  final out = <Map<String, String>>[];
  if (r is List) {
    for (final e in r) {
      if (e is Map) out.add({for (final en in e.entries) en.key.toString(): en.value.toString()});
    }
  }
  return out;
}

// لو نفس التطبيق ظهر بأكتر من طريقة فتح، بنسيب الطريقة الأولى (الشغالة) وبنشيل الباقي (طريقة 2 مش شغالة)
void _numberVariants(List<Map<String, String>> hs) {
  final seen = <String>{};
  hs.removeWhere((h) => !seen.add(h['pkg'] ?? ''));
}

// بيعرض كل التطبيقات المتاحة وانت تختار. لو اخترت "الافتراضي" بيتحفظ وبعدها بيفتح علطول (وتقدر تمسحه من القايمة)
Future<Map<String, String>?> chooseHandler(
  BuildContext context, {
  required String title,
  required List<Map<String, String>> handlers,
  required String prefKey,
  Map<String, String>? systemOption,
}) async {
  final saved = _loadDefaultHandler(prefKey);
  if (saved != null) {
    if (systemOption != null && saved['kind'] == systemOption['kind'] && (saved['pkg'] ?? '').isEmpty) {
      return saved;
    }
    if (handlers.any((h) => h['pkg'] == saved['pkg'] && h['cls'] == saved['cls'] && h['kind'] == saved['kind'])) {
      return saved;
    }
  }
  if (!context.mounted) return null;
  var makeDefault = false;
  return showModalBottomSheet<Map<String, String>>(
    context: context,
    isScrollControlled: true,
    builder: (c) => StatefulBuilder(
      builder: (c, setS) => SafeArea(
        child: ListView(
          shrinkWrap: true,
          children: [
            ListTile(
              title: Text(title, style: const TextStyle(fontWeight: FontWeight.bold)),
            ),
            CheckboxListTile(
              value: makeDefault,
              controlAffinity: ListTileControlAffinity.leading,
              onChanged: (v) => setS(() => makeDefault = v ?? false),
              title: const Text('خليه الافتراضي (يفتح علطول المرة الجاية)'),
            ),
            for (final h in handlers)
              ListTile(
                leading: h['icon'] != null && h['icon']!.isNotEmpty
                    ? Image.memory(base64Decode(h['icon']!), width: 32, height: 32)
                    : const Icon(Icons.apps, size: 32),
                title: Text(h['label'] ?? h['pkg'] ?? '', style: const TextStyle(fontWeight: FontWeight.bold)),
                onTap: () {
                  if (makeDefault) _saveDefaultHandler(prefKey, h);
                  Navigator.pop(c, h);
                },
              ),
            if (systemOption != null)
              ListTile(
                leading: const Icon(Icons.folder_special, size: 32),
                title: Text(systemOption['label'] ?? 'النظام', style: const TextStyle(fontWeight: FontWeight.bold)),
                onTap: () {
                  if (makeDefault) _saveDefaultHandler(prefKey, systemOption);
                  Navigator.pop(c, systemOption);
                },
              ),
          ],
        ),
      ),
    ),
  );
}

// بيجيب كل التطبيقات اللي تقدر تفتح فولدر (مدير ملفات) ويخلّيك تختار أي واحد (وتحفظه افتراضي لو عايز)
Future<bool> chooseAndOpenFolder(BuildContext context, String path) async {
  var handlers = <Map<String, String>>[];
  try {
    handlers = _toHandlers(await _chan.invokeMethod<List<dynamic>>('folderHandlers', {'path': path}));
  } catch (_) {}
  handlers = handlers.where((h) => h['pkg'] != 'com.termux').toList();
  _numberVariants(handlers);
  if (!context.mounted) return false;
  final picked = await chooseHandler(
    context,
    title: 'افتح الفولدر بواسطة',
    handlers: handlers,
    prefKey: 'defFolder',
    systemOption: <String, String>{'pkg': '', 'cls': '', 'kind': 'tree', 'label': 'متصفح ملفات النظام'},
  );
  if (picked == null) return true; // المستخدم قفل القايمة
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

// اختيار صور/ملفات من أي تطبيق معرض أو ملفات (انت اللي تختار وتحفظ الافتراضي).
// بترجع null لو الطريقة دي مش متاحة (نرجع لمنتقي الملفات العادي)، ولستة فاضية لو لغيت.
Future<List<String>?> pickFilesViaChooser(
  BuildContext context, {
  required String mime,
  required String title,
  required String prefKey,
}) async {
  try {
    var handlers = _toHandlers(await _chan.invokeMethod<List<dynamic>>('pickHandlers', {'mime': mime}));
    _numberVariants(handlers);
    if (!context.mounted) return null;
    final picked = await chooseHandler(
      context,
      title: title,
      handlers: handlers,
      prefKey: prefKey,
      systemOption: <String, String>{'pkg': '', 'cls': '', 'kind': 'system', 'label': 'منتقي ملفات النظام'},
    );
    if (picked == null) return <String>[];
    final r = await _chan.invokeMethod<List<dynamic>>('pickWith', {
      'mime': mime,
      'pkg': picked['pkg'] ?? '',
      'cls': picked['cls'] ?? '',
    });
    if (r == null) return null;
    return [for (final e in r) e.toString()];
  } catch (e) {
    _dbg('pickFilesViaChooser error: $e');
    return null;
  }
}

// ───────────────────── الشاشة الرئيسية ─────────────────────
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
  bool autoRotate = false; // أبطأ: بيجرّب تدوير كل صورة مش إيصال
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
    // أول ما التطبيق يفتح: اطلب إذن الملفات (لو مش ممنوح) وانزّل شيت الإعدادات وشيت النتيجة تلقائي
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
      status = 'تم مسح كل حاجة. اضغط "بدء المعالجة الشاملة" لبدء يوم جديد.';
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
              content: w.Text(_l(context, 'كمّل المبلغ والماكينة من زر القلم الأول.',
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

  // سحب إيصال المراجعة: صح ✓ / غلط ✗ (غلط = استبعاد لو مكرر، وإلا بيفتح التصحيح)
  Widget _swipeWrap(ReceiptItem item, Widget card) {
    if (item.ignoredDoc || item.userExcluded || !item.needsReview) return card;
    final okLabel = item.dupSuspect
        ? _l(context, 'احتسبه باتش ✓', 'Count it ✓')
        : (item.dateOdd ? _l(context, 'سيب تاريخه ✓', 'Keep date ✓') : _l(context, 'صح ✓', 'OK ✓'));
    final badLabel = item.dupSuspect
        ? _l(context, 'مكرر - استبعده ✗', 'Duplicate ✗')
        : (item.dateOdd ? _l(context, 'تاريخ الأغلب ✗', 'Use main date ✗') : _l(context, 'غلط - صحّحه ✗', 'Wrong - fix ✗'));
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
      setState(() => status = 'لازم تسمح بالوصول لكل الملفات الأول.');
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
      setState(() => status = 'لازم تسمح بالوصول لكل الملفات الأول.');
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
        setState(() => status = 'مش قادر أكتب في الفولدر ده: $res');
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
      var paths = await pickFilesViaChooser(context,
          mime: 'image/*', title: 'اختر تطبيق المعرض', prefKey: 'defImage');
      if (paths == null) {
        final result = await FilePicker.pickFiles(type: FileType.image);
        if (result.isNotEmpty) {
          paths = result.map((f) => f.path).whereType<String>().toList();
        }
      }
      if (paths != null && paths.isNotEmpty) {
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
      SnackBar(content: Text('مقدرتش أفتح الفولدر. المسار اتنسخ: $path')),
    );
  }

  // إرسال الإكسيل والـ PDF (المستخدم بيختار واتساب من قايمة المشاركة)
  Future<void> sendToWhatsApp() async {
    final x = latestOutput('.xlsx');
    final p = latestOutput('.pdf');
    final files = <XFile>[if (x != null) XFile(x), if (p != null) XFile(p)];
    if (files.isEmpty) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('لم يتم إنشاء ملفات بعد لمشاركتها.')),
      );
      return;
    }
    if (!await confirmSend(context,
        review: processedItems.where((e) => e.needsReview).length, missing: missingTerminals.length)) {
      return;
    }
    await Share.shareXFiles(files, text: 'تقرير الإيصالات');
  }

  // أحدث ملف إخراج (إكسيل/PDF): من آخر معالجة، أو آخر ملف "نتيجة ..." في فولدر الإخراج
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
        const SnackBar(content: Text('الملف مش موجود. شغّل المعالجة الأول.')),
      );
      return;
    }
    try {
      final ok = await _chan.invokeMethod<bool>('openFile', {'path': path, 'mime': mime});
      if (ok == true) return;
    } catch (_) {}
    if (!mounted) return;
    ScaffoldMessenger.of(context).showSnackBar(
      const SnackBar(content: Text('مفيش برنامج يفتح الملف ده.')),
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
          TextButton(onPressed: () => Navigator.pop(ctx, false), child: const Text('إلغاء')),
          FilledButton(onPressed: () => Navigator.pop(ctx, true), child: const Text('مسح')),
        ],
      ),
    );
    return r == true;
  }

  // مسح الكل: ملفات فولدر الإخراج + الصور المرتبة
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
      setState(() => status = 'مفيش حاجة للمسح.');
      return;
    }
    final ok = await _confirm(
      'مسح الكل',
      'هتتمسح ${outFiles.length} ملف من الإخراج:\n$outputDir\n\n'
          'و ${rawFiles.length} صورة من الصور المرتبة.'
          '${isProtected ? '\n\n(فولدر الإخراج ده عام، فهيتمسح بس ملفات التطبيق منه)' : ''}',
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
      status = 'تم مسح $n ملف.';
    });
  }

  Future<void> openReview() async {
    final list = processedItems.where((e) => e.needsReview).toList();
    if (list.isEmpty) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('مفيش إيصالات محتاجة مراجعة.')),
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

  // القلم: شاشة كاملة فيها الإيصال بزوم + تعديل المبلغ والماكينة
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
      await Share.shareXFiles(files, text: 'تقرير الإيصالات (الرسمي)');
    } else {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('لم يتم إنشاء ملفات بعد لمشاركتها.')),
      );
    }
  }

  Future<void> shareAdvancedFile() async {
    final files = <XFile>[];
    for (final p in ['$outputDir/review_report.txt', '$outputDir/ocr_debug.txt']) {
      if (File(p).existsSync()) files.add(XFile(p));
    }
    if (files.isNotEmpty) {
      await Share.shareXFiles(files, text: 'لوج التقفيل (تقرير المراجعة + OCR)');
    } else {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('ملفات المراجعة غير متاحة حالياً.')),
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
          title: const Text('تعديل بيانات الإيصال'),
          content: SingleChildScrollView(
            child: Column(
              mainAxisSize: MainAxisSize.min,
              children: [
                Image.file(File(item.path), height: 160, fit: BoxFit.contain),
                const SizedBox(height: 8),
                TextField(
                  controller: amtController,
                  keyboardType: const TextInputType.numberWithOptions(decimal: true),
                  decoration: InputDecoration(labelText: tr('المبلغ الإجمالي')),
                ),
                const SizedBox(height: 8),
                DropdownButtonFormField<int>(
                  value: selectedOrder,
                  hint: const Text('اختر الماكينة'),
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
            TextButton(onPressed: () => Navigator.pop(ctx), child: const Text('إلغاء')),
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
              child: const Text('حفظ وتحديث الملفات'),
            ),
          ],
        ),
      ),
    );
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

    // 1. ملفات الـ PDF المجمعة: واحد ليوم التقفيل، وواحد لكل يوم تاني لو فيه إيصالات بتاريخ مختلف
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
        setState(() => status = 'جاري تجهيز صفحات الـ PDF ($pageNo / ${processedItems.length})...');
        // سكانر: قص الباتش بس + فرد المنظور + تحسين. لو فشل نستخدم الصورة الأصلية.
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

    // حساب إجماليات كل وحدة
    final unitTotals = <int, double>{};
    for (final r in rowNames.keys) {
      unitTotals[r] = 0.0;
    }
    for (final item in mainItems) {
      if (item.row != null && item.amount > 0) {
        unitTotals[item.row!] = (unitTotals[item.row!] ?? 0.0) + item.amount;
      }
    }

    // الإكسيل: نفس القالب بالظبط، بنكتب بس عمود D (مبلغ كل وحدة) و F (الإجمالي)
    var xlsxLine = '• تم توليد الإكسيل: نتيجة $reportDate.xlsx';
    standardXlsxPath = null;
    final tpl = await loadTemplateBytes();
    if (tpl == null) {
      xlsxLine = '⚠️ قالب شيت النتيجة (template.xlsx) مش موجود في فولدر التطبيق (ولا مضاف للتطبيق)، فمتعملش إكسيل.';
    } else {
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

    setState(() => status = 'اكتمل!\n'
        '$xlsxLine\n'
        '• تم توليد ملف الـ PDF المجمع.' +
        (extraPdfPaths.isEmpty
            ? ''
            : '\n• فيه إيصالات بتواريخ تانية (${otherByDate.keys.join(' ، ')}): اتعمل لها PDF منفصل لكل يوم، ومش داخلة في إجماليات يوم $reportDate.') +
        (sourceNote.isEmpty ? '' : '\n$sourceNote'));
  }

  String subtitleOf(ReceiptItem item) {
    final sfx = item.suffix != null ? ' (باتش ${item.suffix})' : '';
    final bt = item.batches.length > 1
        ? '\n= ${item.batches.map((b) => b.toStringAsFixed(2)).join(' + ')}'
        : '';
    final alt = item.alt != null ? '\n⚠️ قراءة تانية: ${item.alt!.toStringAsFixed(2)}' : '';
    final extra = StringBuffer();
    if (item.ignoredDoc) {
      extra.write('\n📄 اتجاهلته: شكله إيصال دفع/مستند مش باتش - راجعه');
    }
    if (item.dateOdd) {
      extra.write('\n📅 تاريخه ${item.date} والأغلب ${item.dateMain}');
    }
    if (!item.ignoredDoc && !item.userExcluded && !item.dupSuspect && item.weakId && item.row != null) {
      extra.write('\n❓ الماكينة/الوحدة اتحددت تلقائياً - اتأكد');
    }
    final unu = item.unusual != null ? '\n📊 ${item.unusual}' : '';
    final dup = item.dupSuspect
        ? '\n⚠️ مشكوك إنه مكرر (نفس مبلغ ماكينة ${item.dupOf ?? ''}) - مش محسوب لحد ما تراجعه'
        : item.userExcluded
            ? '\n🚫 مستبعد (مكرر)'
            : '';
    return 'TID: ${item.tid ?? 'غير محدد'}$sfx | المبلغ: ${item.amount.toStringAsFixed(2)} ج.م$bt$alt$dup$unu$extra';
  }

  String warningsText() {
    final sb = StringBuffer('ملاحظات للمراجعة (${warnings.length}):');
    for (final w in warnings) {
      sb.write('\n• $w');
    }
    return sb.toString();
  }

  void writeReport(String date) {
    try {
      final line = '=' * 50;
      final sb = StringBuffer('تقرير المراجعة - $date\n$line\n');
      var total = 0.0;
      for (final r in processedItems) {
        final flags = <String>[];
        if (r.dupSuspect) flags.add('مشكوك إنه مكرر - مش محسوب لحد ما تراجعه');
        if (r.userExcluded) flags.add('مستبعد');
        if (r.ignoredDoc) flags.add('صورة متجاهلة - محتاجة قرار');
        if (r.dateOdd) flags.add('تاريخ مختلف ${r.date}');
        if (r.unusual != null) flags.add('إجمالي بعيد عن المعتاد');
        if (r.amount <= 0) flags.add('المبلغ غير مقروء');
        if (r.row == null) flags.add('الوحدة غير معروفة');
        if (!r.ok && r.amount > 0) flags.add('المبلغ يحتاج مراجعة');
        if (r.how == 'auto-number' || r.how == 'conflict' || r.how == 'extra-batch') {
          flags.add('رقم الماكينة: ${r.how}');
        }
        if (r.batches.length > 1) flags.add('${r.batches.length} باتشات');
        if (r.alt != null) flags.add('قراءة تانية: ${r.alt!.toStringAsFixed(2)}');
        if (r.amount > 0 && r.row != null && r.counts) total += r.amount;
        final mark = flags.isEmpty ? '✓' : '⚠️';
        final extra = flags.isEmpty ? '' : '   [${flags.join(' | ')}]';
        sb.writeln('$mark ${baseName(r.path)}  ${r.unitName ?? '??'}  '
            '${r.amount.toStringAsFixed(2)}$extra');
      }
      sb.writeln(line);
      sb.writeln('الإجمالي العام: ${total.toStringAsFixed(2)} ج.م');
      if (missingTerminals.isNotEmpty) {
        sb.writeln('ماكينات ناقصة: ${missingTerminals.join(' ، ')}');
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
          {'title': 'A2R POS', 'text': 'جاري فحص الإيصالات...'});
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
        final txt = 'تم فحص ${processedItems.length} إيصال'
            '${missingTerminals.isNotEmpty ? ' - ناقص ${missingTerminals.length} ماكينة' : ''}'
            '${rv > 0 ? ' - $rv للمراجعة' : ''}';
        try {
          await _chan.invokeMethod('notify', {'title': tr('A2R POS - انتهى الفحص'), 'text': tr(txt)});
        } catch (_) {}
      }
    }
  }

  Future<void> _runPipelineInner({List<String>? manualPaths}) async {
    final perm = await Permission.manageExternalStorage.request();
    if (!perm.isGranted) {
      setState(() => status = 'يجب منح إذن الوصول إلى جميع الملفات للمتابعة.');
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
      sourceNote = '• المصدر: صور محددة يدوياً (${todo.length} صورة)';
    }

    // فولدر اخترته بإيدك (مش واتساب): أي ملف PDF جواه بيتفحص، بأي اسم وبأي تاريخ
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
        setState(() => status = 'بقرأ صفحات ملف PDF: ${baseName(pf.path)}...');
        final pages = await extractPdfPages(pf);
        if (pages.isEmpty) continue;
        todo.addAll(pages);
        pdfPages += pages.length;
        used.add(baseName(pf.path));
      }
      if (used.isNotEmpty) {
        sourceNote = '• المصدر: ملف PDF ${used.take(3).join(' ، ')}${used.length > 3 ? ' ...' : ''} ($pdfPages صفحة)';
      }
    }

    // مفيش صور؟ دوّر على PDF تقفيل بتاريخ النهارده
    if (todo.isEmpty) {
      setState(() => status = 'مفيش صور. بدوّر على ملف PDF تقفيل بتاريخ ${todayToken()}...');
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
          sourceNote = '• المصدر: ملف PDF ${baseName(pdfs.first.path)} (${pages.length} صفحة)';
        }
      }
    }

    if (todo.isEmpty) {
      setState(() {
        busy = false;
        status = customDir != null
            ? 'مفيش صور (في الفترة المحددة) ولا ملفات PDF في المجلد المختار.'
            : (dirs.isEmpty
                    ? 'تعذر العثور على مجلد صور الواتساب تلقائياً.'
                    : 'لا توجد صور في مجلد الواتساب بالفترة المحددة.') +
                '\nولا يوجد ملف PDF تقفيل بتاريخ ${todayToken()}.';
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

    // ═════════ المرحلة 1: قراءة الصور ═════════
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
        if (ign == 'keep') isReceipt = true; // قررت قبل كده إنها باتش
        // إيصال دفع أو مستند مش باتش/ملخص/ري برنت: بيدخل المراجعة (مش بيتتجاهل نهائي)
        var skipPay = false;
        if (isReceipt && ign != 'keep' && classifyReceipt(res.text) == 'payment') {
          isReceipt = false;
          skipPay = true;
        }

        // صورة مقلوبة؟ جرّب EXIF ثم 90 / 270 / 180
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

          // قراءة تانية بتنسيق مختلف (صورة محسّنة) علشان نكشف الأخطاء
          double? alt;
          // القراءة التانية بس للإيصالات اللي القراءة الأولى فيها مش مؤكدة (علشان السرعة)
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

          // نسخة الصورة اللي هتدخل الـ PDF (متظبطة الاتجاه)
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
            await fixImageAsync(f.path, dst, 0); // نسخ + تصحيح EXIF لو لازم
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

          // تصحيح يدوي محفوظ من مرة سابقة
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
        warnings.add('${baseName(f.path)}: فشلت القراءة (${e.toString().split('\n').first})');
      }

      if (!mounted) return;
      setState(() {
        progress = ((i + 1) / todo.length) * 0.5;
        status = 'فحص OCR: ${i + 1} من ${todo.length}';
      });
    }

    await recognizer.close();

    if (rawReceipts.isEmpty) {
      processedItems.addAll(ignoredItems);
      setState(() {
        busy = false;
        status = 'تم فحص ${todo.length} صورة، ولم يتم العثور على إيصالات مطابقة.';
      });
      return;
    }

    // ═════════ المرحلة 2: التكرار وترقيم الماكينات ═════════
    setState(() => status = 'مراجعة التكرار وترقيم الماكينات...');
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
          r.userExcluded = true; // انت قررت قبل كده إنه مكرر (بيفضل ظاهر وتقدر ترجّعه)
          r.dupOf = n;
          warnings.add('${baseName(r.path)}: إيصال مكرر للماكينة $n (استبعدته قبل كده - تقدر ترجّعه)');
        } else if (r.dupDecision == 'keep') {
          r.origTid = r.tid; // انت قررت قبل كده إنه باتش تاني حقيقي
          r.tid = null;
          r.conflict = true;
        } else {
          // نفس الماكينة ونفس المبلغ = مشكوك إنه مكرر: بيدخل المراجعة (مش بيتستبعد)
          r.dupSuspect = true;
          r.dupOf = n;
          warnings.add('${baseName(r.path)}: نفس الماكينة $n ونفس المبلغ '
              '${r.amount.toStringAsFixed(2)} (مشكوك إنه مكرر) - راجعه من زر المراجعة');
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
        giveSuffix(r); // نفس الماكينة بمبلغ مختلف = باتش إضافي
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

    // التاريخ الأكثر تكراراً + تحذيرات
    final dateVotes = <String, int>{};
    for (final r in kept) {
      if (r.date != null) dateVotes[r.date!] = (dateVotes[r.date!] ?? 0) + 1;
    }
    final mainDate = mostCommon(dateVotes.keys.expand((k) => List.filled(dateVotes[k]!, k)).toList())?.key;

    for (final r in kept) {
      final nm = baseName(r.path);
      if (mainDate != null && r.date != null && r.date != mainDate && r.dateFix == null) {
        r.dateOdd = true; // بيدخل المراجعة
        r.dateMain = mainDate;
      }
      if (r.row != null && r.amount > 0) {
        if (!r.ok) warnings.add('$nm: ${r.unitName} = ${r.amount.toStringAsFixed(2)} (راجع المبلغ يدوياً)');
        if (r.alt != null) {
          warnings.add('$nm: القراءة التانية للمبلغ ${r.alt!.toStringAsFixed(2)} مختلفة عن ${r.amount.toStringAsFixed(2)} - راجع الصورة');
        }
        if (r.how == 'extra-batch') {
          warnings.add('$nm: الماكينة ${r.orderNum} ظهرت في صورة تانية بمبلغ مختلف - اتحسبت كباتش إضافي ${r.suffix}');
        } else if (r.how == 'auto-number' || r.how == 'conflict') {
          warnings.add('$nm: رقم الماكينة اتحدد تلقائياً (${r.orderNum}) - اتأكد منه');
        } else if (!(r.how == 'terminal' || r.how == 'fuzzy' || r.how == 'manual')) {
          warnings.add('$nm: الوحدة اتحددت بالاسم/الميرشانت فقط');
        }
      } else {
        warnings.add('$nm: مش قادر أقرا المبلغ أو الوحدة - اضغط تعديل');
      }
    }

    // ═════════ المرحلة 3: إعادة التسمية (على مرحلتين عشان التعارض) ═════════
    setState(() => status = 'إعادة تسمية الصور وترتيبها...');
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
            tooltip: tr('المزيد'),
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
                      const SnackBar(content: Text('مقدرتش أعمل ملف الإعدادات. اتأكد من إذن الوصول لكل الملفات.')),
                    );
                  }
                });
              } else if (v == 'resetdef') {
                resetDefaultHandlers(context);
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
                    child: w.Text(AppSettings.of(context).isAr ? 'المظهر واللغة' : 'Theme & language')),
              const PopupMenuItem(value: 'folder', child: Text('فتح فولدر الإخراج')),
              if (customDir != null)
                PopupMenuItem(value: 'src', enabled: !busy, child: const Text('المصدر: رجوع لواتساب')),
              if (customOut != null)
                PopupMenuItem(value: 'out', enabled: !busy, child: const Text('الإخراج: الافتراضي')),
              PopupMenuItem(
                  value: 'log',
                  child: w.Text(AppSettings.of(context).isAr ? 'مشاركة اللوج' : 'Share log')),
              const PopupMenuItem(enabled: false, child: Text('الإصدار $kAppVersion')),
              PopupMenuItem(value: 'settings', child: const Text('إعدادات الماكينات (إكسيل)')),
              PopupMenuItem(value: 'clear', enabled: !busy, child: const Text('مسح الكل')),
              const PopupMenuItem(value: 'resetdef', child: Text('مسح التطبيقات الافتراضية (فولدر/معرض)')),
              PopupMenuItem(
                  value: 'wipe',
                  enabled: !busy,
                  child: w.Text(AppSettings.of(context).isAr ? '🧹 مسح كل حاجة (كاش + مؤقت)' : '🧹 Wipe everything (cache + temp)')),
            ],
          ),
        ],
      ),
      bottomNavigationBar: _bottomBar(
        FilledButton.icon(
          onPressed: busy ? null : runFullPipeline,
          icon: const Icon(Icons.play_arrow),
          label: const Text('بدء المعالجة الشاملة'),
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
            _pathRow(Icons.folder_open, 'المصدر', customDir ?? 'واتساب (تلقائي)',
                busy ? null : pickFolder,
                onReset: customDir != null && !busy ? useWhatsApp : null),
            _pathRow(Icons.add_photo_alternate, 'اختيار صور يدوياً', 'فتح معرض الصور', busy ? null : pickManualImages),
            _pathRow(Icons.drive_file_move_outline, 'الإخراج',
                customOut ?? '$defaultOutputDir (افتراضي)', busy ? null : pickOutFolder,
                onReset: customOut != null && !busy ? useDefaultOut : null),
            _switchRow('فحص الصور المقلوبة (أبطأ)', autoRotate,
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
                    _l(context, '💡 اسحب إيصال المراجعة: صح ✓ في ناحية وغلط ✗ في الناحية التانية',
                        '💡 Swipe a review receipt: OK ✓ one way, wrong ✗ the other'),
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
                  'إكسيل',
                  xlsx == null
                      ? null
                      : () => openFile(xlsx,
                          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
                  filled: true),
              _mini(Icons.picture_as_pdf, 'PDF',
                  pdf == null ? null : () => openFile(pdf, 'application/pdf'),
                  filled: true),
              _mini(Icons.chat, 'واتساب', busy ? null : sendToWhatsApp),
              _mini(Icons.fact_check, 'مراجعة ($reviewCount)', busy ? null : openReview),
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
                    '⚠️ ماكينات ناقصة (${missingTerminals.length} من ${terminals.length}):',
                    style: TextStyle(
                      fontWeight: FontWeight.bold,
                      color: _tintFg(context, Colors.amber),
                    ),
                  ),
                  const SizedBox(height: 4),
                  Text(
                    missingTerminals
                        .map((n) => '$n (${terminals[numToTid[n]]?.unitName ?? ''})')
                        .join(' ، '),
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
                      title: 'ماكينة ${item.orderNum ?? '؟'} - ${item.unitName ?? ''}',
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
                      ? 'مستند متجاهل - ${baseName(item.path)}'
                      : 'ماكينة ${item.orderNum ?? '؟'} - ${item.unitName ?? 'وحدة غير معروفة'}',
                  style: const TextStyle(fontWeight: FontWeight.bold),
                ),
                subtitle: Text(subtitleOf(item)),
                trailing: Row(
                  mainAxisSize: MainAxisSize.min,
                  children: [
                    if (item.userExcluded)
                      IconButton(
                        tooltip: tr('رجّعه (احتسبه)'),
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
      error = 'مش قادر أفتح الفولدر ده.';
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
      appBar: AppBar(title: const Text('اختيار فولدر')),
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
// ───────── صورة بتتكبّر (قرصة بصباعين + دبل تاب) علشان تشوف تفاصيل الباتش ─────────
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
  final List<ReceiptItem>? allItems; // كل الإيصالات (لحساب رقم الباتش الإضافي)
  final bool single; // فتحناها من القلم لإيصال واحد
  const ReviewPage({super.key, required this.items, this.allItems, this.single = false});

  @override
  State<ReviewPage> createState() => _ReviewPageState();
}

class _ReviewPageState extends State<ReviewPage> {
  int idx = 0;
  bool changed = false;
  int rev = 0; // بيتغير لما الصورة تتدوّر علشان تتحدّث
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
        const SnackBar(content: Text('اكتب مبلغ صحيح.')),
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

  // صورة متجاهلة: ادخلها كباتش، أو تجاهلها نهائي
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

  // تاريخ مختلف: سيبه بتاريخه، أو انقله لتاريخ الأغلب
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

  // قرار الإيصال المشكوك إنه مكرر: استبعده، أو احتسبه كباتش إضافي
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
        const SnackBar(content: Text('مقدرتش أدوّر الصورة.')),
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
              tooltip: tr('ملء الشاشة'),
              icon: const Icon(Icons.fullscreen),
              onPressed: () => Navigator.push(
                context,
                MaterialPageRoute(
                  builder: (_) => ReceiptZoomPage(
                    path: it.path,
                    title: 'ماكينة ${it.orderNum ?? '؟'} - ${it.unitName ?? ''}',
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
                        ? '🚫 الإيصال ده مستبعد (مكرر).'
                        : '⚠️ مشكوك إنه مكرر: نفس الماكينة ${it.orderNum ?? ''} ونفس المبلغ. '
                            'مش محسوب في الإجمالي لحد ما تقرر. زوّم على الباتش وقارن.',
                    style: TextStyle(color: _tintFg(context, Colors.orange), fontSize: 13),
                  ),
                  const SizedBox(height: 6),
                  Row(
                    children: [
                      Expanded(
                        child: FilledButton(
                          onPressed: () => resolveDup(true),
                          child: const Text('مش مكرر - احتسبه باتش إضافي',
                              textAlign: TextAlign.center, style: TextStyle(fontSize: 12)),
                        ),
                      ),
                      if (!it.userExcluded) ...[
                        const SizedBox(width: 6),
                        Expanded(
                          child: OutlinedButton(
                            onPressed: () => resolveDup(false),
                            child: const Text('مكرر - استبعده',
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
                    '📄 التطبيق اتجاهل الصورة دي لأن شكلها إيصال دفع أو مستند مش باتش. '
                    'زوّم وشوف: لو باتش حقيقي ادخله وكمّل المبلغ والماكينة.',
                    style: TextStyle(color: _tintFg(context, Colors.orange), fontSize: 13),
                  ),
                  const SizedBox(height: 6),
                  Row(
                    children: [
                      Expanded(
                        child: FilledButton(
                          onPressed: () => resolveIgnored(true),
                          child: const Text('ده باتش - ادخله',
                              textAlign: TextAlign.center, style: TextStyle(fontSize: 12)),
                        ),
                      ),
                      const SizedBox(width: 6),
                      Expanded(
                        child: OutlinedButton(
                          onPressed: () => resolveIgnored(false),
                          child: const Text('مش باتش - تجاهله',
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
                    '📅 تاريخ الإيصال ${it.date} بينما أغلب الإيصالات ${it.dateMain}.',
                    style: TextStyle(color: _tintFg(context, Colors.orange), fontSize: 13),
                  ),
                  const SizedBox(height: 6),
                  Row(
                    children: [
                      Expanded(
                        child: FilledButton(
                          onPressed: () => resolveDate(true),
                          child: Text('خليه بتاريخ ${it.dateMain}',
                              textAlign: TextAlign.center, style: const TextStyle(fontSize: 12)),
                        ),
                      ),
                      const SizedBox(width: 6),
                      Expanded(
                        child: OutlinedButton(
                          onPressed: () => resolveDate(false),
                          child: const Text('التاريخ صح - سيبه',
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
                    ? '❓ رقم الماكينة اتحدد تلقائياً (${it.orderNum}) - اتأكد منه واضغط حفظ.'
                    : '❓ الوحدة اتحددت بالاسم/الميرشانت فقط - اتأكد من الماكينة واضغط حفظ.',
                style: TextStyle(color: _tintFg(context, Colors.orange), fontSize: 13),
              ),
            ),
          if (it.unusual != null)
            Padding(
              padding: const EdgeInsets.only(bottom: 8),
              child: Text('⚠️ ${it.unusual}  (لو المبلغ صح اضغط حفظ)',
                  style: TextStyle(color: _tintFg(context, Colors.orange), fontSize: 13)),
            ),
          Text(
            'المقروء: ${it.amount > 0 ? it.amount.toStringAsFixed(2) : 'غير مقروء'} ج.م'
            '${it.batches.length > 1 ? '  (= ${it.batches.map((b) => b.toStringAsFixed(2)).join(' + ')})' : ''}',
            style: const TextStyle(fontWeight: FontWeight.bold),
          ),
          if (it.alt != null)
            Padding(
              padding: const EdgeInsets.only(top: 4),
              child: Row(
                children: [
                  Expanded(
                    child: Text('⚠️ القراءة التانية: ${it.alt!.toStringAsFixed(2)}',
                        style: TextStyle(color: _tintFg(context, Colors.orange))),
                  ),
                  TextButton(
                    onPressed: () => setState(() => amt.text = it.alt!.toStringAsFixed(2)),
                    child: const Text('استخدمها'),
                  ),
                ],
              ),
            ),
          const SizedBox(height: 8),
          TextField(
            controller: amt,
            keyboardType: const TextInputType.numberWithOptions(decimal: true),
            decoration: InputDecoration(
              labelText: tr('المبلغ الإجمالي'),
              border: const OutlineInputBorder(),
            ),
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
          Row(
            children: [
              OutlinedButton.icon(
                onPressed: rotate,
                icon: const Icon(Icons.rotate_right),
                label: const Text('تدوير'),
              ),
              const SizedBox(width: 8),
              Expanded(
                child: FilledButton(
                  onPressed: () => save(),
                  child: Text(idx < widget.items.length - 1 ? 'حفظ والتالي' : 'حفظ'),
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
          title: Text('مراجعة ${idx + 1} / ${widget.items.length}'
              '${done.contains(it) ? '  ✓' : ''}'),
          leading: IconButton(
            icon: const Icon(Icons.arrow_back),
            onPressed: () => Navigator.pop(context, changed),
          ),
          actions: [
            IconButton(
              icon: const Icon(Icons.navigate_before),
              tooltip: tr('السابق'),
              onPressed: idx > 0 ? () => go(idx - 1) : null,
            ),
            IconButton(
              icon: const Icon(Icons.navigate_next),
              tooltip: tr('التالي'),
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


// ───────────────────── ملفات الماكينات (PDF لكل ماكينة) ─────────────────────
String? loadPdfDir() => _readCfg()['pdfDir'] as String?;

// فولدرات مستندات واتساب (جنب فولدرات الصور)
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

// الاسم الجديد: "1.اسم.الوحدة.1.يوم.4.11" (وحدة فيها أكتر من ماكينة) أو "17.اسم.الوحدة.يوم.4.11"
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

// ───────── عدّ إيصالات الدفع: كل إيصال دفع ليه سطر RRN (12 رقم، بيبدأ بـ 63) ─────────
// بنقرا الصفحة على شرائح رأسية ضيقة (القراءة أدق من الصفحة كلها)، وبنجمع "نقاط ارتكاز"
// (سطر RRN) لكل إيصال. إيصالات الإقفال والملخص ما فيهاش RRN فمبتتعدّش.
final RegExp _dateRe = RegExp(r'(20\d{2})[^\d\n]{1,3}(\d{2})[^\d\n]{1,3}(\d{2})');

String? _dayKey(String text) {
  for (final m in _dateRe.allMatches(text)) {
    final mo = int.parse(m.group(2)!), d = int.parse(m.group(3)!);
    if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31) return '$d.$mo';
  }
  return null;
}

// null = مفيش RRN في السطر | 12 رقم = RRN كامل | 'P' = RRN مقروء ناقص
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

// بيقص الصفحة لـ 5 شرائح رأسية (أعمدة) متداخلة ويكبّرها. بيرجّع [عرض، ارتفاع، x0 لكل شريحة...]
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
      // نفس الإيصال اتقرا من شريحتين متداخلتين: نفس المكان على الصفحة
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

    // الماكينة: أكتر Terminal ID متكرر
    final tc = <String, int>{};
    for (final m in RegExp(r'(?<!\d)(\d{8})(?!\d)').allMatches(all)) {
      var c = m.group(1)!;
      c = tidFix[c] ?? c;
      if (terminals.containsKey(c)) tc[c] = (tc[c] ?? 0) + 1;
    }
    r.tid = tc.isNotEmpty
        ? tc.entries.reduce((a, b) => a.value >= b.value ? a : b).key
        : (findTerminalIn([all]) ?? discoverTerminal([all]));

    // اليوم الأغلب + الإيصالات اللي مش تبع نفس اليوم
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
        if (d != null && d != best) r.otherDay.add('إيصال ${rrn.substring(6)} بتاريخ $d');
      });
      for (final d in partialDays) {
        if (d != best) r.otherDay.add('إيصال (رقمه مش واضح) بتاريخ $d');
      }
    }

    // تسلسل أرقام الإيصالات (آخر 6 أرقام من الـ RRN): أكبر مجموعة متتالية
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
      seqMin == null ? '' : ' • التسلسل ${seqMin.toString().padLeft(6, '0')}-${seqMax.toString().padLeft(6, '0')} ($rangeLen)';
  String get seqNote {
    final miss = seqMissing.take(10).map((n) => n.toString().padLeft(6, '0')).join('، ');
    return 'العدد $count بس التسلسل $rangeLen'
        '${seqMissing.isNotEmpty ? ' - أرقام مش لاقيها: $miss${seqMissing.length > 10 ? ' ...' : ''}' : ''}'
        ' (ممكن إيصالات ناقصة أو مش اتقرت، راجع الملف)';
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

// قراءة الصفحة كلها + قراءة شرائح مكبّرة، وبعدين دمج النتايج (نفس الإيصال بيتعدّ مرة)
int _tmStrips = 0, _tmOcr = 0; // تشخيص: مجموع الوقت (ms) لعمل الشرائح وللـ OCR

Future<void> ocrPageStrips(
    TextRecognizer rec, String pngPath, String prefix, MCollector col) async {
  final sw = Stopwatch()..start();
  // تقطيع الشرائح (Dart صرف وبطيء) بيشتغل في isolate بالتوازي مع OCR الصفحة الكاملة
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

// ───────── حالة فحص إيصالات الدفع: بتتحفظ بعد كل ملف علشان لو التطبيق اتقفل يكمّل من مكانه ─────────
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
  int count = -1; // -1 = لسه ما اتعدّش
  int tries = 0;
  String detail = '';
  final List<String> idNotes = [];
  final List<String> cntNotes = [];
  List<String> altTids = []; // مرشحين تانيين للماكينة (مش بيتحفظوا)
  Map<String, double> votes = {};
  String? tidUsed, dayUsed, monthUsed;
  bool idOk = true;
  bool cntOk = true;
  int verdict = 0; // 0 = لسه ما اتراجعش، 1 = صح (يفضل مع ملفات اليوم)، -1 = غلط (مستبعد)
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

// تعريف سريع: الماكينة (Terminal ID) + اليوم من نص صفحة أو أكتر
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

// ───────── تعريف الماكينة من جوه الـ PDF بس (تصويت بين الصفحات) — مفيش اعتماد على اسم الملف ─────────
// أصوات صفحة واحدة: "TERMINAL ID" بوزن 3، أي رقم 8 خانات معروف بوزن 1، وتقريبي (غلطة رقم) بوزن 0.5
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

// ───────────────────── تقرير إيصالات الدفع ─────────────────────
String _l(BuildContext c, String ar, String en) => AppSettings.of(c).isAr ? ar : en;

class _NoteKind {
  final String key;
  final int sev; // 2 = لازم تراجع، 1 = انتبه، 0 = معلومة
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
      'العدد مش مطابق لتسلسل الإيصالات',
      'Count does not match the receipt sequence',
      'أرقام الإيصالات (RRN) متتالية، فالعدد المتوقع هو من أول رقم لآخر رقم. العدد اللي اتعدّ مختلف عنه، يعني غالباً فيه إيصال ناقص في الملف أو التطبيق ما قراش إيصال.',
      'Receipt numbers (RRN) are consecutive, so the expected count is first-to-last number. The counted number differs, so a receipt is probably missing from the file or was not read.',
      'افتح الملف وراجع الأرقام اللي مش لاقيها (مكتوبة تحت)، وسجّل العدد الصح في خانة "الصح" لو اتأكدت.',
      'Open the file and check the missing numbers listed below, then enter the right count in the "Correct" field if you are sure.'),
  _NoteKind(
      'otherday',
      2,
      Icons.event_busy,
      'فيه إيصالات من يوم تاني',
      'Receipts from another day',
      'تاريخ بعض الإيصالات مختلف عن يوم باقي الملف، فممكن تكون اتضافت بالغلط أو الملف فيه أكتر من يوم.',
      'Some receipts have a different date from the rest of the file; they may have been added by mistake or the file spans more than one day.',
      'راجع التواريخ المذكورة، ولو الملف مش بتاع اليوم علّمه غلط ✗.',
      'Check the dates listed; if the file is not for this day, mark it wrong ✗.'),
  _NoteKind(
      'machconf',
      2,
      Icons.compare_arrows,
      'تعارض في رقم الماكينة',
      'Machine mismatch',
      'اسم الملف اتحدد لماكينة، لكن الإيصالات جواه بتقول إنها ماكينة تانية.',
      'The file was named for one machine, but the receipts inside belong to another machine.',
      'استخدم "تحديد الماكينة" من القايمة واختار الماكينة الصح.',
      'Use "Assign machine" from the menu and pick the right machine.'),
  _NoteKind(
      'readfail',
      2,
      Icons.error_outline,
      'فشل في قراءة صفحات',
      'Pages failed to read',
      'التطبيق ما قدرش يقرا صفحة أو أكتر (أو اتقفل أثناء العد)، فالعدد ممكن يكون ناقص.',
      'The app could not read one or more pages (or it closed during counting), so the count may be incomplete.',
      'اضغط "إعادة العد" وبعدين "بدء العد". لو اتكرر راجع الملف بنفسك.',
      'Tap "Recount" then "Start counting". If it repeats, check the file yourself.'),
  _NoteKind(
      'emptypages',
      1,
      Icons.insert_drive_file_outlined,
      'صفحات من غير إيصالات',
      'Pages without receipts',
      'فيه صفحات في الملف التطبيق ما لقاش فيها أي إيصال (ممكن تكون صفحة فاضية أو صورة مش واضحة).',
      'Some pages had no receipts detected (they may be blank or unclear).',
      'افتح الصفحات المذكورة وتأكد إنها فعلاً فاضية.',
      'Open the listed pages and make sure they are really empty.'),
  _NoteKind(
      'methoddiff',
      1,
      Icons.difference_outlined,
      'تضارب بين طريقتين للعد',
      'Two counting methods disagree',
      'التطبيق بيعد بطريقتين للتأكد، وفي صفحات معينة النتيجة اختلفت.',
      'The app counts with two methods to double-check, and on some pages the results differ.',
      'راجع الصفحات المذكورة لو العدد مهم ليك.',
      'Check the listed pages if the exact count matters.'),
  _NoteKind(
      'nonpay',
      0,
      Icons.info_outline,
      'مستند إقفال أو ملخص',
      'Closing or summary document',
      'الملف فيه إيصال إقفال أو ملخص أو إعادة طباعة. التطبيق تجاهله ومش محسوب في العدد.',
      'The file contains a closing, summary or reprint receipt. The app ignored it and it is not counted.',
      'مفيش حاجة مطلوبة، ده طبيعي.',
      'Nothing to do, this is normal.'),
  _NoteKind(
      'nofile',
      2,
      Icons.folder_off_outlined,
      'الملف الأصلي مش موجود',
      'Original file not found',
      'الملف اتمسح أو اتنقل من مكانه، فالتطبيق مش قادر يعدّه.',
      'The file was deleted or moved, so the app cannot count it.',
      'رجّع الملف مكانه أو اعمل فحص من جديد.',
      'Put the file back or scan again.'),
  _NoteKind(
      'datefail',
      1,
      Icons.calendar_month_outlined,
      'التاريخ مش واضح',
      'Date not clear',
      'التطبيق ما قدرش يقرا تاريخ اليوم من الملف.',
      'The app could not read the date from the file.',
      'افتح الملف وتأكد من التاريخ بنفسك.',
      'Open the file and check the date yourself.'),
  _NoteKind(
      'namefail',
      2,
      Icons.drive_file_rename_outline,
      'فشلت التسمية',
      'Naming failed',
      'التطبيق ما قدرش يتعرف على الماكينة فمسمّاش الملف.',
      'The app could not identify the machine so it did not rename the file.',
      'استخدم "تحديد الماكينة" من القايمة.',
      'Use "Assign machine" from the menu.'),
  _NoteKind(
      'idweak',
      2,
      Icons.help_outline,
      'الماكينة مش مؤكدة',
      'Machine not confirmed',
      'رقم الماكينة (Terminal ID) ما اتقراش بوضوح، فالتطبيق خمّن الماكينة من بيانات تانية أو لقاها في أكتر من ملف.',
      'The Terminal ID was not read clearly, so the app guessed the machine from other data or found it in more than one file.',
      'اتأكد من الماكينة، ولو غلط استخدم "تحديد الماكينة".',
      'Verify the machine; if wrong, use "Assign machine".'),
  _NoteKind(
      'other',
      1,
      Icons.report_problem_outlined,
      'ملاحظة تانية',
      'Other note',
      'ملاحظة من التطبيق مش متصنفة.',
      'An unclassified note from the app.',
      'اقراها وقرر لو محتاجة مراجعة.',
      'Read it and decide whether it needs review.'),
];

final List<List<Object>> _noteRules = [
  [RegExp(r'^العدد \d+ بس التسلسل'), 'seq'],
  [RegExp(r'مش تبع نفس اليوم'), 'otherday'],
  [RegExp(r'العد لقى ماكينة'), 'machconf'],
  [RegExp(r'صفحة فشلت قراءتها|فشل العد|العد وقف'), 'readfail'],
  [RegExp(r'صفحات ما لقيتش فيها إيصالات'), 'emptypages'],
  [RegExp(r'الطريقتين اختلفوا'), 'methoddiff'],
  [RegExp(r'إقفال/ملخص'), 'nonpay'],
  [RegExp(r'الملف الأصلي مش موجود'), 'nofile'],
  [RegExp(r'مش قادر أعرف التاريخ'), 'datefail'],
  [RegExp(r'فشل التسمية|فشل التعريف'), 'namefail'],
  [RegExp(r'Terminal ID|مش واضح|مش قادر أعرف الماكينة|نفس ماكينة|اتعرفت كمان|الوحدة اتحددت'), 'idweak'],
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

  s.pages = g(RegExp(r'(\d+)\s*صفحة'));
  s.full = g(RegExp(r'RRN\s*(?:مقروء|كامل)\s*(\d+)'));
  s.partial = g(RegExp(r'ناقص القراءة\s*(\d+)'));
  final m = RegExp(r'التسلسل\s*(\d+)-(\d+)\s*\((\d+)\)').firstMatch(d);
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
  if (r.excluded) return const _St('مستبعد', 'Excluded', Colors.grey);
  if (r.pending) return const _St('لسه ما اتعدّش', 'Not counted yet', Colors.blueGrey);
  if (r.verdict > 0) return const _St('اتراجع: صح', 'Reviewed: OK', Colors.green);
  if (r.ok) return const _St('سليم', 'OK', Colors.green);
  return const _St('محتاج مراجعة', 'Needs review', Colors.orange);
}

// الباتش = العدد اللي مكتوب في خانة "الصح" (من إيصال الباتش)، والفعلي = العدد اللي التطبيق عدّه
int? _batchManual(MResult r) => int.tryParse(r.ctl.text.trim());
int? _batchAutoOf(MResult r) => r.orderNum == null ? null : closingTxn.value[r.orderNum!];
int? _batchOf(MResult r) => _batchManual(r) ?? _batchAutoOf(r);
bool _batchIsAuto(MResult r) => _batchManual(r) == null && _batchAutoOf(r) != null;

// مطابقة الباتش: لو انت راجعت الملف يدوي وعلّمته صح ✓ يتحسب مطابق (أخضر) حتى لو العدد المقروء مختلف عن رقم الباتش.
// null = مفيش حاجة نقارنها. الرقم اللي بتكتبه بإيدك في خانة "الصح" بيغلب الرقم التلقائي من التقفيل.
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
    ? _l(c, 'مطابق ✓', 'Match ✓')
    : (d > 0 ? _l(c, 'ناقص $d', 'Missing $d') : _l(c, 'زيادة ${-d}', 'Extra ${-d}'));

MaterialColor _diffColor(int d) => d == 0 ? Colors.green : (d > 0 ? Colors.red : Colors.orange);

String _mSummary(BuildContext c, MResult r) {
  if (r.excluded) return _l(c, 'مستبعد (علامة غلط)', 'Excluded (marked wrong)');
  if (r.pending) return _l(c, 'لسه ما اتعدّش (اضغط بدء العد)', 'Not counted yet (tap Start counting)');
  final b = StringBuffer('${_l(c, 'عدد المعاملات', 'Transactions')}: ${r.count}');
  final bt = _batchOf(r);
  final d = _batchDiff(r);
  if (bt != null && d != null) {
    final tag = _batchIsAuto(r)
        ? _l(c, ' (من التقفيل)', ' (from closing)')
        : (_batchManual(r) != null ? _l(c, ' (يدوي)', ' (manual)') : '');
    final res = (r.verdict > 0 && d != 0)
        ? _l(c, 'راجعته يدوي: صح ✓', 'Manually reviewed: OK ✓')
        : _diffText(c, d);
    b.write('\n${_l(c, 'الباتش', 'Batch')}: $bt$tag • ${_l(c, 'الفعلي', 'Actual')}: ${r.count} • $res');
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
      kv(_l(context, 'عدد المعاملات المعدودة', 'Counted transactions'), r.pending ? '—' : '${r.count}'),
      if (stats.pages != null) kv(_l(context, 'عدد الصفحات', 'Pages'), '${stats.pages}'),
      if (stats.full != null) kv(_l(context, 'إيصالات مقروءة كاملة', 'Fully read receipts'), '${stats.full}'),
      if (stats.partial != null)
        kv(_l(context, 'إيصالات ناقصة القراءة', 'Partially read receipts'), '${stats.partial}',
            color: stats.partial! > 0 ? Colors.orange : null),
      if (stats.seqMin != null)
        kv(_l(context, 'تسلسل الإيصالات', 'Receipt sequence'), '${pad6(stats.seqMin)} → ${pad6(stats.seqMax)}'),
      if (stats.seqLen != null)
        kv(_l(context, 'العدد المتوقع من التسلسل', 'Expected from sequence'), '${stats.seqLen}'),
      if (stats.seqLen != null && !r.pending)
        kv(
            _l(context, 'الفرق عن التسلسل', 'Difference from sequence'),
            r.count == stats.seqLen ? _l(context, 'مطابق ✓', 'Match ✓') : '${r.count - stats.seqLen!}',
            color: r.count == stats.seqLen ? Colors.green : Colors.orange),
    ];

    return Scaffold(
      appBar: AppBar(title: w.Text(_l(context, 'تقرير الإيصال', 'Receipt report'))),
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
                        child: Text(r.newName ?? '؟ ${r.pdfName}',
                            style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 16)),
                      ),
                      const SizedBox(width: 8),
                      _statusChip(context, st),
                    ],
                  ),
                  const SizedBox(height: 4),
                  w.Text('${_l(context, 'الملف الأصلي', 'Original file')}: ${r.pdfName}',
                      style: TextStyle(fontSize: 12, color: cs.onSurfaceVariant)),
                ],
              ),
            ),
          ),
          section(_l(context, 'الأرقام', 'Numbers'), [
            ...numbers,
            if (r.detail.isNotEmpty)
              Padding(
                padding: const EdgeInsets.only(top: 6),
                child: Text(r.detail, style: TextStyle(fontSize: 11, color: cs.onSurfaceVariant)),
              ),
          ]),
          section(_l(context, 'ربط الباتش', 'Batch link'), [
            TextField(
              controller: r.ctl,
              keyboardType: TextInputType.number,
              decoration: InputDecoration(
                isDense: true,
                labelText: _l(context, 'عدد معاملات الباتش (بيتقري من التقفيل، اكتب رقم لو عايز تغيّره)',
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
                  _l(context, 'مفيش عدد باتش اتقرا من التقفيل لماكينة دي. اكتبه هنا علشان نقارنه بالعدد الفعلي.',
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
                    kv(_l(context, 'الباتش', 'Batch'),
                        '$bt${_batchIsAuto(r) ? _l(context, ' (من التقفيل)', ' (from closing)') : ''}'),
                    kv(_l(context, 'الفعلي', 'Actual'), '${r.count}'),
                    kv(
                        _l(context, 'الفرق', 'Difference'),
                        '${_diffText(context, bd)}${r.verdict > 0 && bd != 0 ? _l(context, ' • راجعته يدوي: صح ✓', ' • manually reviewed: OK ✓') : ''}',
                        color: _tintFg(context, r.verdict > 0 ? Colors.green : _diffColor(bd))),
                  ],
                ),
              ),
          ]),
          section(_l(context, 'الملاحظات', 'Notes'), [
            if (kinds.isEmpty)
              Container(
                padding: const EdgeInsets.all(10),
                decoration: BoxDecoration(
                  color: _tintBg(context, Colors.green, soft: true),
                  border: Border.all(color: _tintBorder(context, Colors.green)),
                  borderRadius: BorderRadius.circular(10),
                ),
                child: w.Text(_l(context, 'مفيش ملاحظات ✓', 'No notes ✓'),
                    style: TextStyle(fontWeight: FontWeight.bold, color: _tintFg(context, Colors.green))),
              ),
            for (final k in kinds)
              Builder(builder: (context) {
                final MaterialColor col = k.sev == 2 ? Colors.red : (k.sev == 1 ? Colors.orange : Colors.blue);
                final sevLabel = k.sev == 2
                    ? _l(context, 'لازم تراجع', 'Must review')
                    : (k.sev == 1 ? _l(context, 'انتبه', 'Heads up') : _l(context, 'معلومة', 'Info'));
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
                          child: Text('• $n', style: const TextStyle(fontSize: 12)),
                        ),
                      const SizedBox(height: 6),
                      w.Text('${_l(context, 'تعمل إيه؟', 'What to do')}: ${_l(context, k.fixAr, k.fixEn)}',
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
              _mini(Icons.picture_as_pdf, _l(context, 'فتح الملف', 'Open file'), widget.onOpenPdf, filled: true),
              _mini(Icons.check, _l(context, 'صح ✓', 'OK ✓'), !can
                  ? null
                  : () {
                      widget.onVerdict(1);
                      setState(() {});
                    }),
              _mini(Icons.close, _l(context, 'غلط ✗', 'Wrong ✗'), !can
                  ? null
                  : () {
                      widget.onVerdict(-1);
                      setState(() {});
                    }),
              if (r.verdict != 0)
                _mini(Icons.undo, _l(context, 'إلغاء العلامة', 'Clear mark'), !can
                    ? null
                    : () {
                        widget.onVerdict(0);
                        setState(() {});
                      }),
              _mini(Icons.refresh, _l(context, 'إعادة العد', 'Recount'), !can
                  ? null
                  : () {
                      widget.onRecount();
                      setState(() {});
                      ScaffoldMessenger.of(context).showSnackBar(SnackBar(
                          content: w.Text(_l(context, 'اتعلّم للعد. ارجع واضغط "بدء العد".',
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
  int filter = 0; // 0 الكل، 1 مراجعة، 2 فروق الباتش

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
      appBar: AppBar(title: w.Text(_l(context, 'تقرير إيصالات الدفع', 'Payment receipts report'))),
      body: ListView(
        padding: const EdgeInsets.all(8),
        children: [
          Wrap(
            spacing: 8,
            runSpacing: 8,
            children: [
              stat(_l(context, 'إجمالي المعاملات', 'Total transactions'), '$totalTx', Colors.blue),
              stat(_l(context, 'ملفات', 'Files'), '${all.length}', Colors.blueGrey),
              stat(_l(context, 'اتعدّت', 'Counted'), '${counted.length}', Colors.green),
              stat(_l(context, 'محتاج مراجعة', 'Need review'), '$reviewN', reviewN > 0 ? Colors.orange : Colors.green),
              stat(_l(context, 'مستبعد', 'Excluded'), '$excl', Colors.grey),
              stat(_l(context, 'لسه ما اتعدّتش', 'Not counted'), '$pend', pend > 0 ? Colors.amber : Colors.green),
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
                  w.Text(_l(context, 'ربط الباتش', 'Batch link'),
                      style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 15)),
                  const SizedBox(height: 6),
                  if (linked.isEmpty)
                    w.Text(
                        _l(context, 'عدد الباتش بيتقري تلقائي من إيصالات تاب التقفيل. لو ما ظهرش اكتبه في خانة "الصح" أو في تقرير الإيصال.',
                            'The batch count is read automatically from the closing tab receipts. If it is missing, type it in the "Correct" field or in the receipt report.'),
                        style: TextStyle(fontSize: 12, color: cs.onSurfaceVariant))
                  else ...[
                    kv(_l(context, 'ماكينات مربوطة', 'Linked machines'), '${linked.length} / ${counted.length}'),
                    kv(_l(context, 'مطابقة', 'Matching'), '$matched', color: Colors.green),
                    kv(_l(context, 'فيها فرق', 'With a difference'), '$diffs',
                        color: diffs > 0 ? Colors.red : Colors.green),
                    kv(_l(context, 'إجمالي الباتشات', 'Batch total'), '$batchSum'),
                    kv(_l(context, 'إجمالي الفعلي', 'Actual total'), '$actualSum'),
                    kv(_l(context, 'صافي الفرق', 'Net difference'), _diffText(context, batchSum - actualSum),
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
                  '⚠️ ${_l(context, 'ماكينات ما وصلش ملفها', 'Machines with no file')} (${widget.missing.length}): ${widget.missing.join(' ، ')}',
                  style: const TextStyle(fontSize: 13)),
            ),
          ],
          const SizedBox(height: 8),
          Wrap(
            spacing: 6,
            children: [
              for (final o in [
                [0, _l(context, 'الكل', 'All')],
                [1, _l(context, 'مراجعة ($reviewN)', 'Review ($reviewN)')],
                [2, _l(context, 'فروق الباتش ($diffs)', 'Batch differences ($diffs)')],
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
                title: Text(r.newName ?? '؟ ${r.pdfName}', style: const TextStyle(fontWeight: FontWeight.bold)),
                subtitle: w.Text(_mSummary(context, r)),
                trailing: _statusChip(context, _statusOf(r)),
              ),
            ),
          if (shown.isEmpty)
            Padding(
              padding: const EdgeInsets.all(16),
              child: Center(child: w.Text(_l(context, 'مفيش ملفات في الفلتر ده.', 'No files in this filter.'))),
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
  String? pdfDir; // null = مستندات واتساب (تلقائي)
  int days = 0;
  bool skipDone = true;
  bool reviewOnly = false;
  bool namingDone = false; // التسمية خلصت (حتى لو العد لسه شغال): زرار الإرسال يتفعل
  bool busy = false;
  bool stopReq = false;
  double progressValue = 0;
  String status = 'اضغط "فحص وإعادة تسمية" لفحص إيصالات الدفع من مستندات واتساب.';
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
          ? 'فيه $pend ملف لسه ما اتعدّش. اضغط \"بدء العد\" لما تحب.'
          : 'آخر نتيجة: ${results.length} ملف.';
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
      status = 'تم مسح كل حاجة. اضغط "فحص وإعادة تسمية" لبدء يوم جديد.';
    });
  }

  void _onClosing() {
    if (mounted) setState(() {});
  }

  Future<void> openPdf(String? path) async {
    if (path == null || !File(path).existsSync()) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('الملف مش موجود.')),
      );
      return;
    }
    try {
      final ok = await _chan.invokeMethod<bool>('openFile', {'path': path, 'mime': 'application/pdf'});
      if (ok == true) return;
    } catch (_) {}
    if (!mounted) return;
    ScaffoldMessenger.of(context).showSnackBar(
      const SnackBar(content: Text('مفيش برنامج يفتح الـ PDF.')),
    );
  }

  Future<void> pickDir() async {
    final perm = await Permission.manageExternalStorage.request();
    if (!perm.isGranted) {
      setState(() => status = 'لازم تسمح بالوصول لكل الملفات الأول.');
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
      setState(() => status = 'لازم تسمح بالوصول لكل الملفات الأول.');
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
        setState(() => status = 'مش قادر أكتب في الفولدر ده: $res');
        return;
      }
      saveMOut(res);
      setState(() => destDir = res);
    }
  }

  
  Future<void> pickManualPdfs() async {
    try {
      var paths = await pickFilesViaChooser(context,
          mime: 'application/pdf', title: 'اختر تطبيق الملفات', prefKey: 'defPdf');
      if (paths == null) {
        final result = await FilePicker.pickFiles(type: FileType.custom, allowedExtensions: ['pdf']);
        if (result.isNotEmpty) {
          paths = result.map((f) => f.path).whereType<String>().toList();
        }
      }
      if (paths != null && paths.isNotEmpty) {
        _run(false, false, manualPaths: paths);
      }
    } catch (e) {
      _dbg('pickManualPdfs error: $e');
    }
  }

  Future<void> openFile(String? path, String mime) async {
    if (path == null || !File(path).existsSync()) {
      ScaffoldMessenger.of(context).showSnackBar(const SnackBar(content: Text('الملف مش موجود.')));
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
      SnackBar(content: Text('مقدرتش أفتح الفولدر. المسار اتنسخ: $destDir')),
    );
  }

  // ───── المشاركة: كل الملفات مرة واحدة، أو ملف ملف ─────
  Future<void> shareAll() async {
    final files = results
        .where((r) => !r.excluded && r.newName != null && r.openPath != null && File(r.openPath!).existsSync())
        .map((r) => XFile(r.openPath!, mimeType: 'application/pdf'))
        .toList();
    if (files.isEmpty) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('مفيش ملفات مسمّاة لمشاركتها.')),
      );
      return;
    }
    if (!await confirmSend(context,
        review: results.where((r) => !r.ok).length,
        missing: missingMachines().length,
        pending: results.where((r) => r.pending).length)) {
      return;
    }
    await Share.shareXFiles(files, text: 'إيصالات الدفع (${files.length} ملف)');
  }

  // صح = يفضل مع ملفات اليوم، غلط = يتشال من الإرسال والعد والفولدر. الضغط تاني بيلغي الاختيار
  void setVerdict(MResult r, int v, {bool toggle = true}) {
    final nv = (toggle && r.verdict == v) ? 0 : v;
    final wasExcluded = r.excluded;
    r.verdict = nv;
    final op = r.openPath;
    try {
      if (nv < 0 && !wasExcluded) {
        // مسح النسخة المتسمية بس (مش الأصلي أبداً)
        if (op != null && op != r.srcPath && op.startsWith(destDir)) File(op).deleteSync();
      } else if (nv >= 0 && wasExcluded && r.newName != null) {
        final src = File(r.srcPath);
        if (op != null && op != r.srcPath && src.existsSync()) {
          if (File(op).existsSync()) {
            r.openPath = r.srcPath; // فيه ملف تاني واخد الاسم ده
          } else {
            Directory(File(op).parent.path).createSync(recursive: true);
            src.copySync(op);
          }
        }
      }
    } catch (_) {}
    saveMState(results);
    setState(() {});
    // علّمته صح بس التطبيق مش عارف ماكينته: اسأل عليها وسمّيه وعدّه
    if (nv > 0 && r.newName == null && r.tidUsed == null && !busy) assignMachine(r);
  }

  // تحديد الماكينة بإيدك لملف مش معروف: بيتسمّى وبيتعدّ وبيدخل في إيصالات اليوم
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
          title: const Text('ده إيصال أنهي ماكينة؟'),
          content: DropdownButton<int>(
            isExpanded: true,
            value: sel,
            hint: const Text('اختار الماكينة'),
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
            TextButton(onPressed: () => Navigator.pop(ctx), child: const Text('إلغاء')),
            FilledButton(
              onPressed: sel == null ? null : () => Navigator.pop(ctx, sel),
              child: const Text('تمام'),
            ),
          ],
        ),
      ),
    );
    if (picked == null) return;
    final tid = numToTid[picked];
    String? day = r.dayUsed, month = r.monthUsed;
    if (day == null || month == null) {
      // التاريخ من اسم الملف (آخر شكل 7.11)، وإلا تاريخ النهاردة
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
      r.idNotes.add('فشل التسمية: ${e.toString().split('\n').first}');
    }
    r.verdict = 1;
    r.cntNotes.clear();
    r.count = -1; // يتعدّ من جديد
    r.tries = 0;
    r.detail = '';
    saveMState(results);
    if (!mounted) return;
    setState(() => status = 'اتسمّى ماكينة $picked. اضغط \"بدء العد\" لعدّ إيصالاته.');
  }

  Future<void> markRestOk() async {
    final list = results.where((r) => !r.ok).toList();
    if (list.isEmpty) return;
    final yes = await showDialog<bool>(
      context: context,
      builder: (ctx) => AlertDialog(
        title: const Text('الباقي صح'),
        content: Text('هتعلّم ${list.length} ملف على إنهم صح؟'),
        actions: [
          TextButton(onPressed: () => Navigator.pop(ctx, false), child: const Text('لا')),
          FilledButton(onPressed: () => Navigator.pop(ctx, true), child: const Text('تمام')),
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
        const SnackBar(content: Text('الملف مش موجود.')),
      );
      return;
    }
    await Share.shareXFiles([XFile(p, mimeType: 'application/pdf')],
        text: r.newName ?? r.pdfName);
  }

  Future<void> clearAll() async {
    if (results.isEmpty) {
      setState(() => status = 'مفيش حاجة للمسح.');
      return;
    }
    final mine = results.where((r) => r.newName != null && r.openPath != null).toList();
    final ok = await showDialog<bool>(
      context: context,
      builder: (ctx) => AlertDialog(
        title: const Text('مسح الكل'),
        content: Text('هيتمسح ${mine.length} ملف مسمّى من:\n$destDir\n\n'
            'وهتتمسح نتيجة الفحص. (ملفات واتساب الأصلية مش هتتمس)'),
        actions: [
          TextButton(onPressed: () => Navigator.pop(ctx, false), child: const Text('إلغاء')),
          FilledButton(onPressed: () => Navigator.pop(ctx, true), child: const Text('مسح')),
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
    setState(() => status = 'تم مسح $n ملف ونتيجة الفحص.');
  }

  // ───── التشغيل ─────
  Future<void> start({bool resumeOnly = false, bool count = false}) async {
    if (busy) return;
    final perm = await Permission.manageExternalStorage.request();
    if (!perm.isGranted) {
      setState(() => status = 'لازم تسمح بالوصول لكل الملفات الأول.');
      return;
    }
    try {
      await Permission.notification.request();
    } catch (_) {}
    try {
      await _chan.invokeMethod(
          'startKeepAlive', {'title': 'A2R POS', 'text': 'جاري فحص إيصالات الدفع...'});
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
            'title': tr('A2R POS - انتهى فحص إيصالات الدفع'),
            'text': tr('${results.length} ملف، $total إيصال')
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
      status = 'جاري البحث عن ملفات PDF...';
    });
    try {
      Directory(destDir).createSync(recursive: true);
      Directory(tmpDir).createSync(recursive: true);
    } catch (_) {}
    final rec = TextRecognizer(script: TextRecognitionScript.latin);
    String? idle;
    try {
      // المرحلة 1 (سريعة): تعريف الوحدة وتسمية كل الملفات
      try {
        final lf = File('$baseDir/id_debug.txt');
        if (!resumeOnly && lf.existsSync() && lf.lengthSync() > 300000) lf.deleteSync();
      } catch (_) {}
      _dbg('PHASE1 start resumeOnly=$resumeOnly');
      namingDone = resumeOnly;
      if (!resumeOnly) idle = await _identifyAll(rec, manualPaths: manualPaths);
      _dbg('PHASE1 done');
      if (mounted) setState(() => namingDone = true);
      // المرحلة 2 (بطيئة): عدّ الإيصالات جوه كل ملف: مش بتشتغل غير لما تضغط "بدء العد"
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
        status = 'اتوقف. فاضل $pend ملف لسه ما اتعدّش، اضغط "فحص" وهيكمّل من مكانه.';
      } else if (results.isEmpty && idle != null) {
        status = idle;
      } else if (!count) {
        status = 'تمت التسمية: ${results.length} ملف، $named اتسمّى واتنسخ.'
            '${pend > 0 ? ' اضغط \"بدء العد\" لعدّ إيصالات $pend ملف.' : ''}';
      } else {
        status = 'اكتمل العد: ${results.length} ملف، $named اتسمّى، إجمالي $total إيصال.'
            '${pend > 0 ? ' (فاضل $pend لسه)' : ''}';
      }
    });
  }

  // بيرجّع رسالة لو مفيش حاجة جديدة يفحصها
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
      if (d.path == destDir || d.path == defaultMDest) continue; // ناتج التطبيق نفسه
      try {
        for (final e in d.listSync().whereType<File>()) {
          final n = baseName(e.path);
          final fd = fileDate(e);
          var isPdf = n.toLowerCase().trim().endsWith('.pdf');
          if (!isPdf && !fd.isBefore(cutoff) && !n.substring(n.lastIndexOf('.') + 1).contains(RegExp(r'[A-Za-z0-9]'))) {
            isPdf = _hasPdfHeader(e); // ملفات واتساب DOC-... من غير امتداد
          }
          if (!isPdf) {
            if (!fd.isBefore(cutoff)) _dbg('SKIP(مش PDF) $n  size=${e.lengthSync()}');
            continue;
          }
          // إخراج التقفيل بتاع التطبيق بس (نتيجة 5.11.pdf)، أي اسم تاني بيتفحص عادي
          if (_isAppReportName(n)) {
            _dbg('SKIP(تقرير التطبيق) $n');
            continue;
          }
          if (fd.isBefore(cutoff)) {
            if (fd.isAfter(cutoff.subtract(const Duration(days: 14)))) {
              _dbg('SKIP(أقدم من الفترة: $fd) $n');
            }
            continue;
          }
          if (have.contains('$n:${e.lengthSync()}') && skipDone) {
            skipped++;
            _dbg('SKIP(اتفحص قبل كده) $n');
            continue;
          }
          pdfs.add(e);
        }
      } catch (_) {}
    }
    pdfs.sort((a, b) => baseName(a.path).compareTo(baseName(b.path)));
    if (manualPaths != null) pdfs.addAll(manualPaths.map((p) => File(p)));
    if (pdfs.isEmpty) {
      if (dirs.isEmpty) return 'مش لاقي فولدر مستندات واتساب. اختار فولدر بإيدك.';
      return skipped > 0
          ? 'كل الملفات في الفترة دي اتفحصت قبل كده ($skipped ملف).'
          : 'مفيش ملفات PDF في الفترة دي.';
    }
    // امسح كل النتايج/النسخ القديمة للملفات اللي هنفحصها دلوقتي *قبل* التسمية،
    // علشان أسماء قديمة متعملش تضارب وتطلّع "_2" من غير ما يكون فيه ملف مكرر.
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
      // لو الملف ده اتفحص قبل كده وإحنا بنعيد: امسح الإدخال القديم ونسخته
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
      setState(() => status = 'تعريف الملفات ${i + 1} / ${pdfs.length}: $name');
      try {
        await _identifyOne(r, f, rec);
      } catch (e, st) {
        r.idOk = false;
        r.idNotes.add('فشل التعريف: ${e.toString().split('\n').first}');
        final ss = st.toString();
        _dbg('ERROR(تعريف) $name size=$size: $e\n${ss.substring(0, ss.length > 600 ? 600 : ss.length)}');
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

  // قراءة صفحات الـ PDF (لحد 6) وتصويت على Terminal ID، وبعدين مطابقتها باسم الملف الأصلي
  Future<void> _identifyOne(MResult r, File f, TextRecognizer rec) async {
    final bytes = await f.readAsBytes();
    final sb = StringBuffer();
    final votes = <String, double>{};
    var id = MIdent();
    for (var pi = 0; pi < 25; pi++) {
      // بعد 6 صفحات لو لقينا ماكينة نوقف، ولو لسه مفيش نكمّل لحد 25 صفحة
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
        if (!_pageEnd(e) || pi == 0) rethrow; // آخر صفحة عدّت = مش خطأ
      }
      if (!got) break;
      id = _identify(sb.toString());
      final rk = _rankVotes(votes);
      // وقف بدري لو الماكينة واضحة (Terminal ID مقروء بوضوح) واليوم معروف
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
      // مفيش Terminal ID مقروء: نحدد الوحدة من Merchant ID/اسم الوحدة جوه الـ PDF،
      // ولو الوحدة فيها ماكينة واحدة بس نستخدمها (وتتعلّم للمراجعة)
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
      // مش واضح إنه إيصال ماكينة (توكيل / ملف تاني): تحت المراجعة، من غير تسمية ولا عد
      r.tidUsed = null;
      r.newName = null;
      if (id.day != null && id.month != null) {
        r.dayUsed = id.day;
        r.monthUsed = id.month;
      }
      r.idOk = false;
      r.idNotes.add('الملف ده مش واضح إنه إيصال ماكينة (قراءة ضعيفة للـ Terminal ID)، راجعه');
      r.count = 0;
      r.cntOk = true;
      r.detail = 'للمراجعة (مش معدود)';
      return;
    }
    await _applyName(r, f, tid, id.day, id.month);
    if (fromUnit && r.newName != null) {
      r.idOk = false;
      r.idNotes.add('Terminal ID مش واضح، الوحدة اتحددت من Merchant ID/الاسم جوه الـ PDF، راجع');
    }
  }

  // لو ملفين اتعرفوا على نفس الماكينة: الأقوى (أكتر أصوات) يفضل، والتاني ياخد مرشحه التاني لو قراءته واضحة
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
          r.idNotes.add('الـ PDF اتعرّف على نفس ماكينة ملف تاني، اتسمّى ماكينة ${terminals[alt]!.orderNum} (المرشح التاني من قراءة الـ PDF)، راجع');
        }
      }
    }
  }

  // بيعلّم الملفات اللي لسه واخدة نفس الماكينة (مفيش حل تلقائي)
  void _flagDuplicates() {
    const tag = 'نفس الماكينة';
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
        final others = e.value.where((x) => !identical(x, r)).map((x) => x.pdfName).join(' ، ');
        r.idOk = false;
        r.idNotes.add('$tag (${info?.orderNum ?? ''} ${info?.unitName ?? ''}) اتعرفت كمان في: $others — راجع');
      }
    }
  }

  Future<void> _applyName(MResult r, File f, String? tid, String? day, String? month) async {
    r.idOk = true;
    r.idNotes.clear();
    // نسخة قديمة متسمية للملف ده: نمسحها قبل ما نسمّي من جديد
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
      r.idNotes.add('مش قادر أعرف الماكينة من الـ PDF');
      return;
    }
    final info = terminals[tid]!;
    r.orderNum = info.orderNum;
    if (day == null || month == null) {
      r.idOk = false;
      r.idNotes.add('مش قادر أعرف التاريخ من الـ PDF');
      return;
    }
    r.dayUsed = day;
    r.monthUsed = month;
    final name = machineNewName(info, day, month);
    var final0 = name;
    // "_2" بس لو فعلاً ملف تاني (من نفس الفحص) واخد نفس الاسم؛ أي ملف قديم في الفولدر بنكتب فوقه
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
    // ملف التطبيق اتقفل أثناء عدّه 3 مرات: نعدّيه بدل ما نلف في دايرة
    for (final r in results) {
      if (r.pending && r.tries >= 3) {
        r.count = 0;
        r.cntOk = false;
        r.cntNotes.add('العد وقف أكتر من مرة (التطبيق اتقفل). من ⋮ اختار إعادة العد');
      }
    }
    final todo = results.where((r) => r.pending).toList();
    for (var i = 0; i < todo.length; i++) {
      if (stopReq) break;
      final r = todo[i];
      if (r.excluded) continue; // اتعلّم غلط وإحنا بنعد
      r.tries++;
      saveMState(results);
      setState(() {
        status = 'عدّ الإيصالات ${i + 1} / ${todo.length}: ${r.newName ?? r.pdfName}';
        progressValue = i / todo.length;
      });
      try {
        await _countOne(r, rec);
      } catch (_) {
        r.count = 0;
        r.cntOk = false;
        r.cntNotes.add('فشل العد');
      }
      r.tries = 0;
      saveMState(results);
      _dbg('COUNT ${i + 1}/${todo.length} ${r.newName ?? r.pdfName} = ${r.count}  RSS=${ProcessInfo.currentRss ~/ 1048576}MB');
      if (!mounted) return;
      setState(() => progressValue = (i + 1) / todo.length);
    }
  }


  // العد السريع: بيقرا صورة كل صفحة مرة واحدة من غير تحويل PNG ولا شرائح. بيرجّع false لو الملف مش ممسوح
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
        if ((sp.sm - sp.mx).abs() >= 2) diff.add('${i + 1} (${sp.sm} أو ${sp.mx})');
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
    _dbg('FAST ${r.pdfName}: ${ranges.length} صفحة، ${sw.elapsedMilliseconds ~/ 1000}ث، عدد $total (بالطريقة التانية $totalMx)، فشل $failed');
    if (total == 0 || failed == ranges.length) return false; // نرجع للطريقة القديمة
    final sc = col.finish();
    r.count = total;
    r.detail = 'عدّ سريع • ${ranges.length} صفحة • RRN مقروء ${sc.full}${sc.rangeText}';
    if (failed > 0) {
      r.cntOk = false;
      r.cntNotes.add('$failed صفحة فشلت قراءتها، العدد ممكن يكون ناقص');
    }
    if (empty.isNotEmpty) {
      r.cntNotes.add('صفحات ما لقيتش فيها إيصالات: ${empty.join('، ')}');
    }
    if (diff.isNotEmpty) {
      r.cntNotes.add('صفحات الطريقتين اختلفوا فيها: ${diff.join('، ')} (راجعها لو العدد مهم)');
    }
    // نقارن بتسلسل الـ RRN بس لو الأرقام اتقرت كفاية
    if (sc.rangeLen != null && sc.full >= total * 0.8 && total != sc.rangeLen) {
      r.cntOk = false;
      final miss = sc.seqMissing.take(10).map((n) => n.toString().padLeft(6, '0')).join('، ');
      r.cntNotes.add('العدد $total بس التسلسل ${sc.rangeLen}'
          '${sc.seqMissing.isNotEmpty ? ' - أرقام مش لاقيها: $miss${sc.seqMissing.length > 10 ? ' ...' : ''}' : ''}');
    }
    if (sc.otherDay.length >= 3 && sc.otherDay.length >= total * 0.1) {
      r.cntOk = false;
      r.cntNotes.add('مش تبع نفس اليوم: ${sc.otherDay.take(5).join(' ، ')}');
    }
    if (sc.nonPayment) {
      r.cntNotes.add('فيه إيصال إقفال/ملخص/إعادة طباعة (اتجاهل ومش معدود)');
    }
    if (r.newName == null && sc.tid != null) {
      await _applyName(r, src, sc.tid, sc.day, sc.month);
    } else if (sc.tid != null && terminals[sc.tid] != null && r.orderNum != null &&
        terminals[sc.tid]!.orderNum != r.orderNum) {
      r.cntOk = false;
      r.cntNotes.add('العد لقى ماكينة ${terminals[sc.tid]!.orderNum} بس التسمية ماكينة ${r.orderNum}، راجع');
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
      r.cntNotes.add('الملف الأصلي مش موجود');
      return;
    }
    final bytes = await src.readAsBytes();
    if (fastCount) {
      try {
        if (await _countOneFast(r, rec, bytes, src)) return;
      } catch (e) {
        _dbg('FAST فشل وهنكمل بالطريقة الدقيقة: $e');
      }
      r.cntNotes.clear();
      r.cntOk = true;
    }
    final col = MCollector();
    _tmStrips = 0;
    _tmOcr = 0;
    var nPages = 0, tRaster = 0;
    final swAll = Stopwatch()..start();
    double? dpiCache; // مقاس الصفحات في الملف الممسوح واحد: نقيس مرة بس
    // صفحة صفحة (مش الملف كله مرة واحدة) علشان الذاكرة ما تخلصش والتطبيق ما يقفلش
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
    _dbg('PAGES ${r.pdfName}: $nPages صفحة، المجموع ${swAll.elapsedMilliseconds ~/ 1000}ث '
        '(raster+png ${tRaster ~/ 1000}ث، شرائح ${_tmStrips ~/ 1000}ث، OCR ${_tmOcr ~/ 1000}ث)');
    final sc = col.finish();
    r.count = sc.count;
    r.detail = 'RRN كامل ${sc.full} • ناقص القراءة ${sc.partial}${sc.rangeText}';
    if (!sc.agree) {
      r.cntOk = false;
      r.cntNotes.add(sc.seqNote);
    }
    if (sc.otherDay.isNotEmpty) {
      r.cntOk = false;
      r.cntNotes.add('مش تبع نفس اليوم: ${sc.otherDay.join(' ، ')}');
    }
    if (sc.nonPayment) {
      r.cntNotes.add('فيه إيصال إقفال/ملخص/إعادة طباعة (اتجاهل ومش معدود)');
    }
    if (r.newName == null && sc.tid != null) {
      // التعريف السريع فشل: نسمّي من نتيجة العد الكاملة
      await _applyName(r, src, sc.tid, sc.day, sc.month);
    } else if (sc.tid != null && terminals[sc.tid] != null && r.orderNum != null &&
        terminals[sc.tid]!.orderNum != r.orderNum) {
      r.cntOk = false;
      r.cntNotes.add('العد لقى ماكينة ${terminals[sc.tid]!.orderNum} بس التسمية ماكينة ${r.orderNum}، راجع');
    }
  }

  Future<void> recount(MResult r) async {
    r.count = -1;
    r.tries = 0;
    saveMState(results);
    if (mounted) setState(() => status = 'اتعلّم للعد. اضغط \"بدء العد\".');
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
        const SnackBar(content: Text('مفيش لوج لسه. شغّل الفحص الأول.')),
      );
      return;
    }
    await Share.shareXFiles([XFile(f.path, mimeType: 'text/plain')], text: 'لوج إيصالات الدفع');
  }

  @override
  Widget build(BuildContext context) {
    return PopScope(
      canPop: !busy,
      onPopInvokedWithResult: (didPop, _) async {
        if (!didPop && busy) {
          // الفحص شغال: ما نقفلش الشاشة، نرجّع التطبيق للخلفية بس علشان الفحص يكمّل
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
            tooltip: tr('المزيد'),
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
                      const SnackBar(content: Text('مقدرتش أعمل ملف الإعدادات. اتأكد من إذن الوصول لكل الملفات.')),
                    );
                  }
                });
              } else if (v == 'resetdef') {
                resetDefaultHandlers(context);
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
                    child: w.Text(AppSettings.of(context).isAr ? 'المظهر واللغة' : 'Theme & language')),
              const PopupMenuItem(value: 'folder', child: Text('فتح فولدر الإخراج')),
              if (pdfDir != null)
                PopupMenuItem(value: 'src', enabled: !busy, child: const Text('المصدر: رجوع لواتساب')),
              if (destDir != defaultMDest)
                PopupMenuItem(value: 'out', enabled: !busy, child: const Text('الإخراج: الافتراضي')),
              PopupMenuItem(
                  value: 'log',
                  child: w.Text(AppSettings.of(context).isAr ? 'مشاركة اللوج' : 'Share log')),
              const PopupMenuItem(enabled: false, child: Text('الإصدار $kAppVersion')),
              PopupMenuItem(value: 'fast', child: Text(fastCount ? 'العد: سريع ← الدقيق (القديم)' : 'العد: دقيق ← السريع')),
              PopupMenuItem(value: 'settings', child: const Text('إعدادات الماكينات (إكسيل)')),
              PopupMenuItem(value: 'clear', enabled: !busy, child: const Text('مسح الكل')),
              const PopupMenuItem(value: 'resetdef', child: Text('مسح التطبيقات الافتراضية (فولدر/معرض)')),
              PopupMenuItem(
                  value: 'wipe',
                  enabled: !busy,
                  child: w.Text(AppSettings.of(context).isAr ? '🧹 مسح كل حاجة (كاش + مؤقت)' : '🧹 Wipe everything (cache + temp)')),
            ],
          ),
        ],
      ),
      bottomNavigationBar: _bottomBar(
        busy
            ? OutlinedButton.icon(
                onPressed: stopReq ? null : () => setState(() => stopReq = true),
                icon: const Icon(Icons.stop),
                label: const Text('إيقاف'),
              )
            : Row(
                children: [
                  Expanded(
                    child: OutlinedButton.icon(
                      onPressed: () => start(),
                      icon: const Icon(Icons.drive_file_rename_outline),
                      label: const Text('فحص وتسمية'),
                    ),
                  ),
                  const SizedBox(width: 8),
                  Expanded(
                    child: FilledButton.icon(
                      onPressed: pendCount > 0 ? () => start(resumeOnly: true, count: true) : null,
                      icon: const Icon(Icons.play_arrow),
                      label: Text(pendCount > 0 ? 'بدء العد ($pendCount)' : 'بدء العد'),
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
            _pathRow(Icons.folder_open, 'المصدر', pdfDir ?? 'واتساب - مستندات (تلقائي)',
                busy ? null : pickDir,
                onReset: pdfDir != null && !busy
                    ? () {
                        savePdfDir(null);
                        setState(() => pdfDir = null);
                      }
                    : null),
            _pathRow(Icons.add_photo_alternate, 'اختيار PDF يدوياً', 'فتح معرض الملفات', busy ? null : pickManualPdfs),
            _pathRow(Icons.drive_file_move_outline, 'الإخراج',
                '$destDir${destDir == defaultMDest ? ' (افتراضي)' : ''}', busy ? null : pickOut,
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
              _mini(Icons.chat, 'إرسال الكل ($named)',
                  (busy && !namingDone) || named == 0 ? null : shareAll,
                  filled: true),
              _mini(Icons.insights, _l(context, 'التقرير', 'Report'), results.isEmpty ? null : openOverview,
                  filled: true),
              if (reviewOnly && reviewN > 0)
                _mini(Icons.done_all, 'الباقي صح ($reviewN)', busy ? null : markRestOk),
            ],
          ),
          Padding(
            padding: const EdgeInsets.only(top: 4, bottom: 2),
            child: Text('اسحب أي ملف لتعليمه: صح ✓ أو غلط ✗',
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
              child: Text('⚠️ ماكينات ما وصلش ملفها (${missing.length}): ${missing.join(' ، ')}',
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
                  child: const Text('صح ✓', style: TextStyle(fontWeight: FontWeight.bold)),
                ),
                secondaryBackground: Container(
                  color: _tintBg(context, Colors.red, strong: true),
                  alignment: AlignmentDirectional.centerEnd,
                  padding: const EdgeInsets.symmetric(horizontal: 20),
                  child: const Text('غلط ✗', style: TextStyle(fontWeight: FontWeight.bold)),
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
                    title: Text(r.newName ?? '؟ ${r.pdfName}',
                        style: TextStyle(
                            fontWeight: FontWeight.bold,
                            decoration: r.excluded ? TextDecoration.lineThrough : null)),
                    subtitle: w.Text(
                        '${r.verdict > 0 ? _l(context, '✓ اتراجع: صح\n', '✓ Reviewed: OK\n') : ''}'
                        '${r.pending && busy ? _l(context, 'بيتعدّ...', 'Counting...') : _mSummary(context, r)}'
                        '${r.notes.isEmpty ? '' : '\n${_l(context, '⚠️ ${r.notes.length} ملاحظة - اضغط لفتح التقرير', '⚠️ ${r.notes.length} note(s) - tap to open the report')}'}'),
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
                                hintText: _batchAutoOf(r)?.toString() ?? tr('الصح'),
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
                            tooltip: tr('المزيد'),
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
                                  child: w.Text(_l(context, 'تقرير الإيصال', 'Receipt report'))),
                              PopupMenuItem(
                                  value: 'open',
                                  child: w.Text(_l(context, 'فتح الملف', 'Open file'))),
                              PopupMenuItem(
                                  value: 'share',
                                  enabled: r.newName != null && !r.excluded,
                                  child: const Text('إرسال الملف')),
                              PopupMenuItem(
                                  value: 'assign', enabled: !busy, child: const Text('تحديد الماكينة')),
                              PopupMenuItem(
                                  value: 'recount', enabled: !busy, child: const Text('إعادة العد')),
                              if (r.verdict != 0)
                                const PopupMenuItem(value: 'clear', child: Text('إلغاء علامة صح/غلط')),
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
