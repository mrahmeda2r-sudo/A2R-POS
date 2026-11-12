// ════════════════════════════════════════════════════════════════════
// المرحلة 1 + 2: سجل الماكينات (المصدر الأساسي) + تعديل التعلّم من الفحص
// الملف ده مش بيتبني لوحده. هو تعليمات لصق على main.dart:
//   (A) كتلة جديدة تتلصق مرة واحدة
//   (B) دوال موجودة بتتبدّل بالكامل أو من نقطة معينة
// ════════════════════════════════════════════════════════════════════


// ════════════════ (A) كتلة جديدة ════════════════
// الصقها بعد سطر:  final Map<int, String> rowNames = {};
// (وسيب rebuildTerminalIndexes زي ما هي، مش هتتنده تاني غير من المستورد القديم)

// سطر واحد = ماكينة واحدة. الترتيب في القايمة = رقم الماكينة (index + 1)
// tid فاضي = مكان محجوز (ماكينة لسه ما اشتغلتش)
class MachineRow {
  String tid;
  String unit; // اسم الوحدة (عربي أو إنجليزي)
  String merchant; // MERCHANT ID: هو اللي بيجمّع الماكينات في إجمالي واحد
  String status; // active | off (إجازة)
  bool auto; // اتعلّمت تلقائي من الفحص
  bool lockName; // الاسم اتعدّل يدوي: الفحص ما يغيّرهوش
  bool lockMerch;
  String flag; // new-unit | no-merchant | by-name  (للمراجعة)
  Map<String, int> votes; // تصويت على الاسم من الإيصالات
  int row = 0; // رقم الوحدة (بيتحسب من الـ Merchant ID، مش بيتحفظ)

  MachineRow({
    this.tid = '',
    this.unit = '',
    this.merchant = '',
    this.status = 'active',
    this.auto = false,
    this.lockName = false,
    this.lockMerch = false,
    this.flag = '',
    Map<String, int>? votes,
  }) : votes = votes ?? <String, int>{};

  bool get reserved => tid.isEmpty;

  Map<String, dynamic> toJson() => {
        'tid': tid,
        'unit': unit,
        'merchant': merchant,
        'status': status,
        'auto': auto,
        'lockName': lockName,
        'lockMerch': lockMerch,
        'flag': flag,
        'votes': votes,
      };

  static MachineRow fromJson(Map<String, dynamic> j) => MachineRow(
        tid: (j['tid'] ?? '').toString(),
        unit: (j['unit'] ?? '').toString(),
        merchant: (j['merchant'] ?? '').toString(),
        status: (j['status'] ?? 'active').toString(),
        auto: j['auto'] == true,
        lockName: j['lockName'] == true,
        lockMerch: j['lockMerch'] == true,
        flag: (j['flag'] ?? '').toString(),
        votes: {
          for (final e in ((j['votes'] as Map?) ?? {}).entries)
            e.key.toString(): (e.value as num).toInt()
        },
      );
}

final List<MachineRow> registry = [];
final String registryPath = '$baseDir/.app_registry.json';
bool _regLoaded = false;
final List<String> registryWarnings = []; // تنبيهات الفحص الحالي (المرحلة 4 هتعرضها)

// تطابق كامل بس. الاستثناء الوحيد: كود أقصر (10 أرقام أو أكتر) هو نهاية كود أطول
// (OCR بيلزّق أرقام زيادة قدّام الكود). فرق رقم واحد في النص = وحدة تانية.
bool merchantEq(String a, String b) {
  if (a.isEmpty || b.isEmpty) return false;
  if (a == b) return true;
  final s = a.length <= b.length ? a : b;
  final l = a.length <= b.length ? b : a;
  return s.length >= 10 && l.endsWith(s);
}

MachineRow? machineByTid(String tid) {
  for (final m in registry) {
    if (m.tid == tid) return m;
  }
  return null;
}

// الماكينات اللي لو ما ظهرتش في اليوم تتنبّه عليها: مش المحجوزة ومش الإجازة
List<int> expectedMachineNums() {
  final out = <int>[];
  for (var i = 0; i < registry.length; i++) {
    final m = registry[i];
    if (!m.reserved && m.status == 'active') out.add(i + 1);
  }
  return out;
}

String _unitNameOfRow(int row) {
  for (final m in registry) {
    if (m.row == row && m.lockName && m.unit.trim().isNotEmpty) return m.unit.trim();
  }
  for (final m in registry) {
    if (m.row == row && m.unit.trim().isNotEmpty) return m.unit.trim();
  }
  return 'وحدة $row';
}

// بيبني الفهارس القديمة من السجل. لازم يتنده بعد أي تغيير في registry
void rebuildFromRegistry() {
  terminals.clear();
  numToTid.clear();
  rowNames.clear();
  merchantCodeMap.clear();
  final reps = <int, MachineRow>{};
  var next = 0;
  for (final m in registry) {
    int? row;
    if (m.merchant.isNotEmpty) {
      for (final e in reps.entries) {
        if (merchantEq(e.value.merchant, m.merchant)) {
          row = e.key;
          break;
        }
      }
    }
    if (row == null) {
      row = ++next;
      reps[row] = m;
    }
    m.row = row;
  }
  for (var i = 0; i < registry.length; i++) {
    final m = registry[i];
    final name = _unitNameOfRow(m.row);
    rowNames.putIfAbsent(m.row, () => name);
    if (m.merchant.isNotEmpty) merchantCodeMap.putIfAbsent(m.merchant, () => m.row);
    if (!m.reserved) {
      terminals[m.tid] = TerminalInfo(i + 1, m.row, name);
      numToTid[i + 1] = m.tid;
    }
  }
}

void saveRegistry() {
  try {
    File(registryPath)
      ..createSync(recursive: true)
      ..writeAsStringSync(jsonEncode(registry.map((m) => m.toJson()).toList()));
  } catch (e) {
    _dbg('registry save error: $e');
  }
}

// أول تشغيل: لو فيه A2R_Settings.xlsx قديم بيتنقل مرة واحدة لجوه السجل. غير كده السجل بيبدأ فاضي.
Future<void> loadRegistry() async {
  if (_regLoaded) return;
  try {
    final f = File(registryPath);
    if (f.existsSync()) {
      final l = jsonDecode(f.readAsStringSync());
      registry.clear();
      if (l is List) {
        for (final e in l.whereType<Map>()) {
          registry.add(MachineRow.fromJson(Map<String, dynamic>.from(e)));
        }
      }
    } else if (File(settingsPath).existsSync()) {
      await loadSettingsFromExcel(force: true);
      registry.clear();
      final byNum = terminals.entries.toList()
        ..sort((a, b) => a.value.orderNum.compareTo(b.value.orderNum));
      for (final e in byNum) {
        var mc = '';
        for (final k in merchantCodeMap.entries) {
          if (k.value == e.value.excelRow) {
            mc = k.key;
            break;
          }
        }
        registry.add(MachineRow(
          tid: e.key,
          unit: e.value.unitName,
          merchant: mc,
          lockName: true,
          lockMerch: mc.isNotEmpty,
        ));
      }
      saveRegistry();
    }
    _regLoaded = true;
  } catch (e) {
    _dbg('registry load error: $e');
  }
  rebuildFromRegistry();
}

// ماكينة جديدة ظهرت في إيصال. القواعد:
//  1) فيه سطر محجوز (TID فاضي) بنفس الـ Merchant ID تمام -> بيملاه (أول واحد بالترتيب)
//  2) فيه سطر محجوز من غير Merchant ID واسمه يشبه الاسم المقروء -> بيملاه + تنبيه "by-name"
//  3) غير كده -> بتتضاف في آخر القايمة. لو الـ Merchant ID مطابق لوحدة موجودة بتدخل إجماليها
//     لوحدها (rebuildFromRegistry بيجمّع بالكود). لو مش مطابق لأي حاجة: وحدة جديدة + تنبيه.
String? registerLearned(String tid, String? name, String? mid) {
  if (machineByTid(tid) != null) return tid;
  final hadRows = registry.isNotEmpty;
  MachineRow? slot;
  var how = '';
  if (mid != null) {
    for (final m in registry) {
      if (m.reserved && m.merchant.isNotEmpty && merchantEq(m.merchant, mid)) {
        slot = m;
        how = 'merchant';
        break;
      }
    }
  }
  if (slot == null && name != null) {
    final k = _normUnit(name);
    for (final m in registry) {
      if (!m.reserved || m.merchant.isNotEmpty || m.unit.isEmpty) continue;
      final o = _normUnit(m.unit);
      final mx = math.max(o.length, k.length);
      if (o == k || (mx >= 6 && _editDistance(o, k) <= (mx * 0.15).floor())) {
        slot = m;
        how = 'name';
        break;
      }
    }
  }
  if (slot != null) {
    slot.tid = tid;
    slot.status = 'active';
    slot.auto = true;
    if (slot.merchant.isEmpty && mid != null) slot.merchant = mid;
    if (slot.unit.isEmpty && name != null) slot.unit = name;
    if (how == 'name') slot.flag = 'by-name';
    _dbg('RESERVED FILLED ($how) $tid -> ماكينة ${registry.indexOf(slot) + 1}');
    registryWarnings.add('ماكينة ${registry.indexOf(slot) + 1} (${slot.unit}) اتملت من مكان محجوز'
        '${how == 'name' ? ' بالاسم فقط - اتأكد' : ''}');
  } else {
    var flag = '';
    if (mid == null) {
      flag = hadRows ? 'no-merchant' : '';
    } else if (hadRows && !registry.any((m) => merchantEq(m.merchant, mid))) {
      flag = 'new-unit';
    }
    slot = MachineRow(tid: tid, unit: name ?? '', merchant: mid ?? '', auto: true, flag: flag);
    registry.add(slot);
    _dbg('NEW TERMINAL $tid -> ${name ?? '?'} merchant=${mid ?? '?'} (ماكينة ${registry.length}) flag=$flag');
    if (flag.isNotEmpty) {
      registryWarnings.add('ماكينة جديدة ${registry.length} ($tid): '
          '${flag == 'new-unit' ? 'Merchant ID مش مطابق لأي وحدة - اتضافت كوحدة جديدة' : 'Merchant ID مش مقروء - مش معروف بتتبع أنهي وحدة'}');
    }
  }
  if (name != null) slot.votes[name] = (slot.votes[name] ?? 0) + 1;
  _learnedThisRun.add(tid);
  rebuildFromRegistry();
  saveRegistry();
  return tid;
}


// ════════════════ (B) دوال بتتبدّل ════════════════

// ── B1. بدّل learnMerchant بالكامل ──
// بتكمّل Merchant ID الناقص، وبتصوّت على الاسم (إلا لو انت قفلته). أي تعارض بيتسجل تنبيه بس، مش بيغيّر حاجة.
void learnMerchant(String tid, List<String> texts) {
  final m = machineByTid(tid);
  if (m == null) return;
  var ch = false;
  final mid = extractMerchantId(texts);
  if (mid != null) {
    if (m.merchant.isEmpty && !m.lockMerch) {
      m.merchant = mid;
      ch = true;
    } else if (m.merchant.isNotEmpty && !merchantEq(m.merchant, mid)) {
      final msg = 'ماكينة ${registry.indexOf(m) + 1} ($tid): الـ Merchant ID المقروء $mid مختلف عن المسجل ${m.merchant}';
      if (!registryWarnings.contains(msg)) registryWarnings.add(msg);
    }
  }
  if (!m.lockName) {
    final nm = extractUnitName(texts);
    if (nm != null) {
      m.votes[nm] = (m.votes[nm] ?? 0) + 1;
      final best = m.votes.entries.reduce((a, b) => a.value >= b.value ? a : b).key;
      if (best != m.unit) {
        m.unit = best;
        ch = true;
      }
    }
  }
  if (ch) {
    rebuildFromRegistry();
    saveRegistry();
  }
}

// ── B2. بدّل _rowForMerchant بالكامل ──
int? _rowForMerchant(String mid) {
  for (final e in merchantCodeMap.entries) {
    if (merchantEq(e.key, mid)) return e.value;
  }
  return null;
}

// ── B3. في _discoverTerminalStrict: امسح كل حاجة بعد السطر
//        `if (best.value < 2) return null;`
//        وحط بداله: ──
//   final mid = extractMerchantId(texts);
//   final uname = extractUnitName(texts);
//   final knownUnit = (mid != null && _rowForMerchant(mid) != null) ||
//       (uname != null && _existingRowForName(uname) != null);
//   if (!knownUnit && best.value < 5) return null;
//   return registerLearned(tid, uname, mid);

// ── B4. في discoverTerminal (غير strict): امسح كل حاجة بعد السطر
//        `if (nk.isNotEmpty) return nk;`
//        (يعني من `final name = ...` لحد `return tid;`) وحط بداله: ──
//   final name = extractUnitName(texts);
//   final mid = extractMerchantId(texts);
//   return registerLearned(tid, name, mid);

// ── B5. saveSettingsToExcel: غيّر اسم الدالة الموجودة لـ exportSettingsToExcel
//        (هتفضل احتياطي للتصدير)، وضيف الدالة دي بالاسم القديم علشان كل النداءات تفضل شغالة ──
bool saveSettingsToExcel() {
  saveRegistry();
  return true;
}

// ── B6. بدّل prepareDefaultFiles بالكامل ──
Future<void> prepareDefaultFiles() async {
  try {
    if (await Permission.manageExternalStorage.isGranted) {
      await ensureDefaultFiles();
      await loadRegistry(); // بعد الإذن بس، علشان ما نعتبرش السجل فاضي غلط
    }
  } catch (_) {}
}

// ── B7. في ensureDefaultFiles: امسح بلوك نسخ A2R_Settings.xlsx من الـ assets
//        (من `final st = File(settingsPath);` لحد قفلة الـ if بتاعه).
//        سيب بلوك template.xlsx زي ما هو. ──

// ── B8. بدّل openSettingsFile (مؤقتاً لحد شاشة الإدارة في المرحلة 3) ──
Future<void> openSettingsFile(BuildContext context) async {
  final ok = exportSettingsToExcel();
  if (!context.mounted) return;
  if (!ok) {
    ScaffoldMessenger.of(context).showSnackBar(
      const SnackBar(content: Text('مقدرتش أعمل ملف الإعدادات. اتأكد من إذن الوصول لكل الملفات.')),
    );
    return;
  }
  try {
    await _chan.invokeMethod<bool>('openFile', {
      'path': settingsPath,
      'mime': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
  } catch (_) {}
}

// ── B9. تنبيه الماكينات الناقصة (استثناء الإجازة والمحجوز) ──
// في HomePage.rebuildFiles: غيّر  `for (final i in machineNums())`  في missingTerminals لـ
//   `for (final i in expectedMachineNums())`
// في _MachinesPageState.missingMachines: غيّر  `for (final n in machineNums())`  لـ
//   `for (final n in expectedMachineNums())`
// (قايمة اختيار الماكينة machineNums() تفضل زي ما هي، علشان تختار حتى ماكينة إجازة)

// ── B10. تصفير تنبيهات الفحص: في بداية _runPipelineInner (جنب _learnedThisRun.clear())
//        وفي بداية _run بتاعة MachinesPage ضيف:  registryWarnings.clear();
