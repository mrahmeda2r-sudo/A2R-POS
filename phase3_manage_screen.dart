// ════════════════════════════════════════════════════════════════════
// المرحلة 3: شاشة إدارة الماكينات
// (A) الصق الكتلة دي كلها في main.dart (آخر الملف مناسب)
// (B) تعديلين صغيرين في _moreMenu (تحت)
// ════════════════════════════════════════════════════════════════════

// ════════════════ (A) ════════════════

Future<void> openMachinesManager(BuildContext context) async {
  if (homeBusyG || machBusyG || chainRunning) {
    ScaffoldMessenger.of(context).showSnackBar(
      const SnackBar(content: Text('استنى لحد ما الفحص يخلص أو وقّفه الأول.')),
    );
    return;
  }
  await Navigator.push(context, MaterialPageRoute(builder: (_) => const MachinesManagePage()));
}

// استيراد احتياطي من ملف A2R_Settings.xlsx: بيستبدل السجل كله
Future<bool> importRegistryFromExcel() async {
  await loadSettingsFromExcel(force: true);
  if (terminals.isEmpty) {
    rebuildFromRegistry(); // القارئ القديم مسح الفهارس: نرجّعها من السجل
    return false;
  }
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
  rebuildFromRegistry();
  saveRegistry();
  return true;
}

String _flagText(String f) => f == 'new-unit'
    ? '⚠️ Merchant ID مش مطابق لأي وحدة - راجعه'
    : f == 'no-merchant'
        ? '⚠️ Merchant ID مش مقروء - مش معروف تبع أنهي وحدة'
        : f == 'by-name'
            ? '⚠️ اتملت من مكان محجوز بالاسم فقط - اتأكد'
            : '';

String? _cleanTid(String raw) {
  var d = raw.replaceAll(RegExp(r'\D'), '');
  if (d.length > 8) d = d.substring(d.length - 8);
  return d.length == 8 ? d : null;
}

class MachinesManagePage extends StatefulWidget {
  const MachinesManagePage({super.key});

  @override
  State<MachinesManagePage> createState() => _MachinesManagePageState();
}

class _MachinesManagePageState extends State<MachinesManagePage> {
  void _commit() {
    rebuildFromRegistry();
    saveRegistry();
    if (mounted) setState(() {});
  }

  void _snack(String m) => ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: w.Text(m)));

  bool _tidTaken(String tid, MachineRow except) => registry.any((x) => !identical(x, except) && x.tid == tid);

  // ───── تعديل سطر ─────
  Future<void> _edit(int i) async {
    final m = registry[i];
    final tid = TextEditingController(text: m.tid);
    final name = TextEditingController(text: m.unit);
    final mer = TextEditingController(text: m.merchant);
    var off = m.status == 'off';
    final ok = await showDialog<bool>(
      context: context,
      builder: (ctx) => StatefulBuilder(
        builder: (ctx, setS) => AlertDialog(
          title: w.Text('ماكينة ${i + 1}'),
          content: SingleChildScrollView(
            child: Column(
              mainAxisSize: MainAxisSize.min,
              children: [
                TextField(
                  controller: tid,
                  keyboardType: TextInputType.number,
                  decoration: InputDecoration(labelText: tr('Terminal ID (فاضي = مكان محجوز)')),
                ),
                const SizedBox(height: 8),
                TextField(
                  controller: name,
                  textDirection: TextDirection.rtl,
                  decoration: InputDecoration(labelText: tr('اسم الوحدة')),
                ),
                const SizedBox(height: 8),
                TextField(
                  controller: mer,
                  keyboardType: TextInputType.number,
                  decoration: InputDecoration(labelText: tr('Merchant ID')),
                ),
                SwitchListTile(
                  contentPadding: EdgeInsets.zero,
                  title: const Text('إجازة (مش بتتحسب ناقصة)'),
                  value: off,
                  onChanged: (v) => setS(() => off = v),
                ),
              ],
            ),
          ),
          actions: [
            TextButton(onPressed: () => Navigator.pop(ctx, false), child: const Text('إلغاء')),
            FilledButton(onPressed: () => Navigator.pop(ctx, true), child: const Text('حفظ')),
          ],
        ),
      ),
    );
    final tRaw = tid.text, nRaw = name.text.trim(), mRaw = mer.text.replaceAll(RegExp(r'\D'), '');
    tid.dispose();
    name.dispose();
    mer.dispose();
    if (ok != true) return;
    String newTid = '';
    if (tRaw.trim().isNotEmpty) {
      final t = _cleanTid(tRaw);
      if (t == null) {
        _snack('Terminal ID لازم 8 أرقام.');
        return;
      }
      if (_tidTaken(t, m)) {
        _snack('الـ Terminal ID ده مسجل في ماكينة تانية.');
        return;
      }
      newTid = t;
    }
    if (nRaw != m.unit) {
      m.unit = nRaw;
      m.lockName = nRaw.isNotEmpty; // اتعدّل يدوي: الفحص ما يغيّرهوش
    }
    if (mRaw != m.merchant) {
      m.merchant = mRaw;
      m.lockMerch = mRaw.isNotEmpty;
    }
    m.tid = newTid;
    m.status = off ? 'off' : 'active';
    m.flag = '';
    m.auto = false;
    _commit();
  }

  Future<void> _delete(int i) async {
    final m = registry[i];
    final ok = await showDialog<bool>(
      context: context,
      builder: (ctx) => AlertDialog(
        title: w.Text('مسح ماكينة ${i + 1}؟'),
        content: const Text('اللي بعدها هيتزحزح رقمه. الأسماء القديمة للصور هتتحدث في الفحص الجاي.'),
        actions: [
          TextButton(onPressed: () => Navigator.pop(ctx, false), child: const Text('إلغاء')),
          FilledButton(
            style: FilledButton.styleFrom(backgroundColor: Colors.red),
            onPressed: () => Navigator.pop(ctx, true),
            child: const Text('مسح'),
          ),
        ],
      ),
    );
    if (ok != true) return;
    registry.remove(m);
    _commit();
  }

  // ───── إضافة ماكينة (محجوزة) في مكان معين ─────
  Future<void> _addAt() async {
    final pos = TextEditingController(text: '${registry.length + 1}');
    final name = TextEditingController();
    final mer = TextEditingController();
    final ok = await showDialog<bool>(
      context: context,
      builder: (ctx) => AlertDialog(
        title: const Text('إضافة ماكينة في مكان معين'),
        content: SingleChildScrollView(
          child: Column(
            mainAxisSize: MainAxisSize.min,
            children: [
              TextField(
                controller: pos,
                keyboardType: TextInputType.number,
                decoration: InputDecoration(labelText: tr('في أنهي رقم؟ (اللي بعده هيتزحزح)')),
              ),
              const SizedBox(height: 8),
              TextField(
                controller: name,
                textDirection: TextDirection.rtl,
                decoration: InputDecoration(labelText: tr('اسم الوحدة')),
              ),
              const SizedBox(height: 8),
              TextField(
                controller: mer,
                keyboardType: TextInputType.number,
                decoration: InputDecoration(labelText: tr('Merchant ID (ده اللي بيربطها بالوحدة)')),
              ),
              const SizedBox(height: 8),
              const Text('هتفضل محجوزة (من غير Terminal ID) لحد ما تظهر في فحص بنفس الـ Merchant ID فتتسجل لوحدها.',
                  style: TextStyle(fontSize: 12)),
            ],
          ),
        ),
        actions: [
          TextButton(onPressed: () => Navigator.pop(ctx, false), child: const Text('إلغاء')),
          FilledButton(onPressed: () => Navigator.pop(ctx, true), child: const Text('إضافة')),
        ],
      ),
    );
    final p = int.tryParse(pos.text.trim()) ?? registry.length + 1;
    final n = name.text.trim();
    final d = mer.text.replaceAll(RegExp(r'\D'), '');
    pos.dispose();
    name.dispose();
    mer.dispose();
    if (ok != true) return;
    if (n.isEmpty && d.isEmpty) {
      _snack('اكتب اسم الوحدة أو الـ Merchant ID.');
      return;
    }
    final at = (p - 1).clamp(0, registry.length);
    registry.insert(at, MachineRow(unit: n, merchant: d, lockName: n.isNotEmpty, lockMerch: d.isNotEmpty));
    _commit();
    _snack('اتضافت ماكينة محجوزة رقم ${at + 1}');
  }

  // ───── لصق عمود واحد (سطر لكل ماكينة بالترتيب) ─────
  Future<void> _pasteColumn(String kind) async {
    final title = kind == 'tid' ? 'لصق Terminal IDs' : kind == 'merchant' ? 'لصق Merchant IDs' : 'لصق أسامي الوحدات';
    final txt = TextEditingController();
    final start = TextEditingController(text: '1');
    final ok = await showDialog<bool>(
      context: context,
      builder: (ctx) => AlertDialog(
        title: Text(title),
        content: SingleChildScrollView(
          child: Column(
            mainAxisSize: MainAxisSize.min,
            children: [
              TextField(
                controller: start,
                keyboardType: TextInputType.number,
                decoration: InputDecoration(labelText: tr('ابدأ من ماكينة رقم')),
              ),
              const SizedBox(height: 8),
              TextField(
                controller: txt,
                maxLines: 8,
                textDirection: kind == 'name' ? TextDirection.rtl : TextDirection.ltr,
                decoration: InputDecoration(
                  hintText: tr('سطر لكل ماكينة'),
                  border: const OutlineInputBorder(),
                ),
              ),
              Align(
                alignment: AlignmentDirectional.centerStart,
                child: TextButton.icon(
                  icon: const Icon(Icons.paste, size: 18),
                  label: const Text('لصق من الحافظة'),
                  onPressed: () async {
                    final d = await Clipboard.getData('text/plain');
                    if (d?.text != null) txt.text = d!.text!;
                  },
                ),
              ),
              const Text('السطر الفاضي بيتخطّى (الماكينة بتفضل زي ما هي). لو العمود أطول من القايمة بيتضاف سطور جديدة.',
                  style: TextStyle(fontSize: 12)),
            ],
          ),
        ),
        actions: [
          TextButton(onPressed: () => Navigator.pop(ctx, false), child: const Text('إلغاء')),
          FilledButton(onPressed: () => Navigator.pop(ctx, true), child: const Text('تطبيق')),
        ],
      ),
    );
    final text = txt.text;
    final s = (int.tryParse(start.text.trim()) ?? 1).clamp(1, registry.length + 1);
    txt.dispose();
    start.dispose();
    if (ok != true) return;
    final lines = text.split(RegExp(r'\r?\n'));
    while (lines.isNotEmpty && lines.last.trim().isEmpty) {
      lines.removeLast();
    }
    var idx = s - 1, done = 0, bad = 0;
    for (final line in lines) {
      final cur = idx++;
      final raw = line.trim();
      if (raw.isEmpty) continue;
      // لو اتنسخ من إكسيل بأكتر من عمود: ناخد أول خلية
      final v = raw.contains('\t') ? raw.split('\t').first.trim() : raw;
      if (v.isEmpty) continue;
      String? val;
      if (kind == 'tid') {
        val = _cleanTid(v);
      } else if (kind == 'merchant') {
        final d = v.replaceAll(RegExp(r'\D'), '');
        val = d.isEmpty ? null : d;
      } else {
        val = v;
      }
      if (val == null) {
        bad++;
        continue;
      }
      while (registry.length <= cur) {
        registry.add(MachineRow());
      }
      final m = registry[cur];
      if (kind == 'tid') {
        if (_tidTaken(val, m)) {
          bad++;
          continue;
        }
        m.tid = val;
        m.auto = false;
      } else if (kind == 'merchant') {
        m.merchant = val;
        m.lockMerch = true;
      } else {
        m.unit = val;
        m.lockName = true;
      }
      m.flag = '';
      done++;
    }
    _commit();
    _snack('اتلصق $done${bad > 0 ? ' • اتخطّى $bad (رقم غلط أو مكرر)' : ''}');
  }

  // ───── إكسيل احتياطي ─────
  Future<void> _exportExcel() async {
    if (!exportSettingsToExcel()) {
      _snack('مقدرتش أعمل الملف.');
      return;
    }
    await Share.shareXFiles([XFile(settingsPath)], text: 'نسخة احتياطية: ماكينات A2R');
  }

  Future<void> _importExcel() async {
    final yes = await showDialog<bool>(
      context: context,
      builder: (ctx) => AlertDialog(
        title: const Text('استيراد من إكسيل'),
        content: const Text('القايمة الحالية كلها هتتستبدل بالموجود في الملف. تكمّل؟'),
        actions: [
          TextButton(onPressed: () => Navigator.pop(ctx, false), child: const Text('لا')),
          FilledButton(onPressed: () => Navigator.pop(ctx, true), child: const Text('كمّل')),
        ],
      ),
    );
    if (yes != true) return;
    try {
      final result = await FilePicker.pickFiles(type: FileType.custom, allowedExtensions: ['xlsx']);
      if (result.isEmpty) return;
      final p = result.first.path;
      if (p == null) return;
      if (p != settingsPath) File(p).copySync(settingsPath);
      final ok = await importRegistryFromExcel();
      _snack(ok ? 'اتستورد ${registry.length} ماكينة.' : 'الملف مفيهوش ماكينات.');
      if (mounted) setState(() {});
    } catch (e) {
      _dbg('import excel error: $e');
      _snack('فشل الاستيراد.');
    }
  }

  @override
  Widget build(BuildContext context) {
    AppSettings.of(context);
    final cs = Theme.of(context).colorScheme;
    final flagged = registry.where((m) => m.flag.isNotEmpty).length;
    final reserved = registry.where((m) => m.reserved).length;
    return Scaffold(
      appBar: AppBar(
        title: const Text('إدارة الماكينات'),
        actions: [
          PopupMenuButton<String>(
            onSelected: (v) {
              if (v == 'export') _exportExcel();
              if (v == 'import') _importExcel();
            },
            itemBuilder: (_) => const [
              PopupMenuItem(value: 'export', child: Text('تصدير نسخة إكسيل (احتياطي)')),
              PopupMenuItem(value: 'import', child: Text('استيراد من إكسيل')),
            ],
          ),
        ],
      ),
      body: SafeArea(
        child: Column(
          children: [
            Padding(
              padding: const EdgeInsets.fromLTRB(8, 8, 8, 0),
              child: Wrap(
                spacing: 6,
                runSpacing: 6,
                children: [
                  _mini(Icons.pin, 'لصق Terminal IDs', () => _pasteColumn('tid')),
                  _mini(Icons.badge_outlined, 'لصق Merchant IDs', () => _pasteColumn('merchant')),
                  _mini(Icons.drive_file_rename_outline, 'لصق الأسامي', () => _pasteColumn('name')),
                  _mini(Icons.add_location_alt_outlined, 'إضافة في مكان معين', _addAt, filled: true),
                ],
              ),
            ),
            Padding(
              padding: const EdgeInsets.fromLTRB(12, 8, 12, 4),
              child: Row(
                children: [
                  Expanded(
                    child: w.Text(
                      '${registry.length} ماكينة • ${rowNames.length} وحدة • $reserved محجوزة'
                      '${flagged > 0 ? ' • $flagged للمراجعة' : ''}',
                      style: TextStyle(fontSize: 12, color: cs.onSurfaceVariant),
                    ),
                  ),
                ],
              ),
            ),
            Expanded(
              child: registry.isEmpty
                  ? const Center(
                      child: Padding(
                        padding: EdgeInsets.all(24),
                        child: Text('القايمة فاضية. اعمل فحص أول أو الصق البيانات من فوق.',
                            textAlign: TextAlign.center),
                      ),
                    )
                  : ReorderableListView.builder(
                      buildDefaultDragHandles: false,
                      padding: const EdgeInsets.fromLTRB(8, 0, 8, 16),
                      itemCount: registry.length,
                      onReorder: (o, n) {
                        if (n > o) n--;
                        final m = registry.removeAt(o);
                        registry.insert(n, m);
                        _commit();
                      },
                      itemBuilder: (_, i) {
                        final m = registry[i];
                        final ft = _flagText(m.flag);
                        return Card(
                          key: ObjectKey(m),
                          margin: const EdgeInsets.symmetric(vertical: 3),
                          shape: RoundedRectangleBorder(
                            borderRadius: BorderRadius.circular(12),
                            side: BorderSide(color: _unitColor(m.row), width: 1.5),
                          ),
                          child: ListTile(
                            onTap: () => _edit(i),
                            leading: CircleAvatar(
                              backgroundColor: _unitColor(m.row),
                              foregroundColor: Colors.white,
                              child: w.Text('${i + 1}'),
                            ),
                            title: w.Text(
                              m.reserved ? tr('محجوزة (لسه من غير Terminal ID)') : m.tid,
                              textDirection: m.reserved ? null : TextDirection.ltr,
                              style: TextStyle(
                                  fontWeight: FontWeight.bold,
                                  color: m.reserved ? cs.onSurfaceVariant : null),
                            ),
                            subtitle: w.Text(
                              '${m.unit.isEmpty ? tr('من غير اسم') : m.unit}'
                              '${m.merchant.isEmpty ? '' : '\nMerchant: ${m.merchant}'}'
                              '${m.status == 'off' ? '\n${tr('🏖 إجازة')}' : ''}'
                              '${ft.isEmpty ? '' : '\n${tr(ft)}'}',
                              style: const TextStyle(fontSize: 12),
                            ),
                            isThreeLine: true,
                            trailing: Row(
                              mainAxisSize: MainAxisSize.min,
                              children: [
                                if (m.flag.isNotEmpty)
                                  IconButton(
                                    tooltip: tr('تمام، سيب التنبيه'),
                                    icon: const Icon(Icons.check_circle_outline, color: Colors.orange),
                                    onPressed: () {
                                      m.flag = '';
                                      _commit();
                                    },
                                  ),
                                IconButton(
                                  tooltip: tr('مسح'),
                                  icon: const Icon(Icons.delete_outline),
                                  onPressed: () => _delete(i),
                                ),
                                ReorderableDragStartListener(
                                  index: i,
                                  child: const Padding(
                                    padding: EdgeInsets.all(6),
                                    child: Icon(Icons.drag_handle),
                                  ),
                                ),
                              ],
                            ),
                          ),
                        );
                      },
                    ),
            ),
          ],
        ),
      ),
    );
  }
}

// ════════════════ (B) تعديل _moreMenu ════════════════
// 1) في onSelected: غيّر
//      } else if (v == 'settings') {
//        openSettingsFile(context);
//    لـ
//      } else if (v == 'settings') {
//        openMachinesManager(context);
// 2) في itemBuilder: غيّر النص
//      PopupMenuItem(value: 'settings', child: w.Text(ar ? 'إعدادات الماكينات (إكسيل)' : 'Machines settings (Excel)')),
//    لـ
//      PopupMenuItem(value: 'settings', child: w.Text(ar ? 'إدارة الماكينات' : 'Manage machines')),
// (دالة openSettingsFile القديمة تقدر تمسحها أو تسيبها)
