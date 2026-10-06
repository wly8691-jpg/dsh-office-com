// kb_annotate / kb_verify — P3-a（工单 2026-10-03 第 12/13 条）
//
// 标注写端：annotations.jsonl 每行是 ContextItem 形状 + 标注字段；登记优先，
// 写回原文件（仅 .md 前置元数据）默认关，要 write_back && confirm 双闸。
// 核账：标注可解析、源文件仍在且指纹吻合、影子方案与磁盘一致。
//
// 自包含模块：STATE_DIR 公式与 index.mjs 相同（process.env.DSH_OFFICE_STATE_DIR
// || ~/.dsh-office-com）；身份哈希与 pool_scan / kb blob 逐字同源
// （sha256(rel_posix \0 size \0 mtime_ns)）。
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readdirSync, readFileSync, realpathSync, statSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'

export const STATE_DIR = process.env.DSH_OFFICE_STATE_DIR || join(homedir(), '.dsh-office-com')

const fingerprintOf = (relPosix, size, mtimeNs) => createHash('sha256')
  .update(`${relPosix}\x00${size}\x00${mtimeNs}`).digest('hex')

// B1（DSH 验收 2026-10-06）：Node 的 fs.Stats **没有** `st_mtime_ns` —— 那是 Python
// `os.stat_result` 的属性名，在 JS 里静默取到 `undefined`，于是这里算出的指纹**永远
// 不等于** registry 的指纹（registry 由 Python 侧 pool_scan 算）。后果：标注恒
// `unverifiable`、kb_verify 恒报 fingerprint-changed、写回后重算的"新指纹"同样错。
// 等价值是 bigint 口径的 `mtimeNs`（实测与 Python st_mtime_ns 逐位相同）。
const statNs = (p) => {
  const st = statSync(p, { bigint: true })
  return { size: st.size, mtimeNs: st.mtimeNs }
}

const rootIdOf = (rootReal) => createHash('sha256').update(rootReal).digest('hex').slice(0, 16)

// W4（工单 2026-10-06 第九批）：Office 写回与 W2 读回共用的扩展名表（写/读两处必须同一口径）。
const OFFICE_EXTS = ['.xlsx', '.xlsm', '.xls', '.docx', '.doc', '.pptx', '.ppt']

export async function kbAnnotateImpl(args, helpers = {}) {
  let rootReal
  try {
    rootReal = realpathSync(args.root)
  } catch {
    return { ok: false, error: '[OPEN_FAILED] 目录不存在: ' + args.root }
  }
  if (!args.fingerprint) {
    return { ok: false, error: '[MISSING_PARAM] 缺少 fingerprint（来自 kb_scan_root 的 registry）' }
  }
  const rootId = rootIdOf(rootReal)
  const dir = join(STATE_DIR, 'kb', rootId)
  const regPath = join(dir, 'registry.json')
  let registry = null
  try {
    registry = JSON.parse(readFileSync(regPath, 'utf-8'))
  } catch {
    return { ok: false, error: '[OPEN_FAILED] 没有 registry——先跑 kb_scan_root' }
  }
  const e = (registry.entries || {})[args.fingerprint]
  const relPosix = e?.source_uri?.split('://')[1] ||
    (args.source_uri ? String(args.source_uri).split('://')[1] : null)
  // 指纹判定（★4 修正 2026-10-05）：这里算的是「标注时刻」的文件状态；
  // 若随后发生写回，写回后会重新 stat 并把**新**指纹存进标注（见下方 fileModified 分支），
  // 否则 kb_verify 会拿旧指纹对已改文件永久报警（OCR ★5）。
  let unverifiable = false
  let fpAtAnnotation = args.fingerprint   // 标注里落的指纹（缺省=调用方给的）
  if (e && relPosix) {
    const abs = join(rootReal, relPosix)
    try {
      const s = statNs(abs)
      fpAtAnnotation = fingerprintOf(relPosix, s.size, s.mtimeNs)
      unverifiable = fpAtAnnotation !== args.fingerprint   // 指纹已变 = 内容已变
    } catch {
      unverifiable = true                              // 文件已不在
      fpAtAnnotation = args.fingerprint
    }
  } else {
    unverifiable = true
  }
  if (unverifiable && args.require_intact === true) {
    return { ok: false, error: '[UNVERIFIABLE] 文件已不在或指纹已变（可用 require_intact=false 只登记）' }
  }
  const writeBack = args.write_back === true && args.confirm === true
  let fileModified = false
  let wbWarning = null
  if (writeBack) {
    if (!relPosix) {
      wbWarning = 'write-back 被拒：没有可定位的原文件'
    } else if (OFFICE_EXTS.some((x) => relPosix.toLowerCase().endsWith(x))) {
      // W4（工单 2026-10-06 第九批）：Office 写回 —— 经 seam 调 kb_write_back_office，
      // **只写 kb_* 自定义文档属性，绝不碰单元格/正文/幻灯片**（红线 1）。
      // 与语义缝的关键差别：**写入类不做 direct 降级** —— 语义缝那种「拿工具对象直接
      // execute」会绕过审批，对写动作不可接受；拿不到审批通道就只登记、不写。
      const abs = join(rootReal, relPosix)
      if (typeof helpers.officeWriter !== 'function') {
        wbWarning = 'write-back 被拒：Office 写回需 host 提供审批 seam（本条只登记）'
      } else {
        const props = {}
        if (args.note) props.kb_note = String(args.note)
        if (args.taxonomy_id) props.kb_taxonomy = String(args.taxonomy_id)
        if (args.sensitivity) props.kb_sensitivity = String(args.sensitivity)
        if (!Object.keys(props).length) {
          wbWarning = 'write-back 被拒：没有可写字段（note / taxonomy_id / sensitivity 全空）'
        } else {
          try {
            const r = await helpers.officeWriter(abs, props)
            if (!r || r.error) throw new Error(String((r && r.error) || 'writer returned nothing'))
            fileModified = true
          } catch (e2) {
            wbWarning = 'Office 写回失败: ' + String(e2)
          }
        }
      }
      if (fileModified) {
        // 与 .md 分支同理（★4/★5）：写回改了 mtime ⇒ 指纹必须重算并存入标注，
        // 否则 kb_verify 拿旧指纹对已改文件永久报警，重扫也不自愈。
        try {
          const s2 = statNs(abs)
          fpAtAnnotation = fingerprintOf(relPosix, s2.size, s2.mtimeNs)
          unverifiable = false
        } catch {
          unverifiable = true
        }
      }
    } else if (!relPosix.toLowerCase().endsWith('.md')) {
      wbWarning = 'write-back 被拒：仅支持 .md 前置元数据与 Office 自定义属性（本条只登记）'
    } else {
      const abs = join(rootReal, relPosix)
      try {
        let text = readFileSync(abs, 'utf-8')
        const noteLine = 'kb_note: ' + String(args.note || '').replace(/[\r\n]+/g, ' ')
        if (text.startsWith('---\n') || text.startsWith('---\r\n')) {
          const nl = text.indexOf('\n') + 1
          const end = text.indexOf('---', nl)
          if (end < 0) throw new Error('前置元数据未闭合')
          text = text.slice(0, end) + noteLine + '\n' + text.slice(end)
        } else {
          text = '---\n' + noteLine + '\n---\n\n' + text
        }
        writeFileSync(abs, text, 'utf-8')
        fileModified = true
      } catch (e2) {
        wbWarning = 'write-back 失败: ' + String(e2)
      }
      if (fileModified) {
        // ★4/★5：写回成功 → 重新 stat，把**写回后**的指纹存进标注——
        // 否则 kb_verify 拿旧指纹对已改文件永久报警，重扫也不会自愈
        try {
          const s2 = statNs(abs)
          fpAtAnnotation = fingerprintOf(relPosix, s2.size, s2.mtimeNs)
          unverifiable = false          // 写回后文件在盘上且完整
        } catch {
          unverifiable = true           // 极端：写完读不到 → 如实标不可验证
        }
      }
    }
  } else if (args.write_back === true) {
    // B6（DSH 验收 2026-10-06）：双闸缺一即不写 —— 但「不写」必须给出说明。
    // 原实现所有 wbWarning 赋值都在 if (writeBack) 之内，缺 confirm 时那句永远轮不到执行，
    // 结果 write_back:false 配 wb_warning:null —— 调用方无从知道是"被闸挡住"还是"没这回事"。
    wbWarning = 'write-back 未执行：双闸需 write_back:true 且 confirm:true（缺 confirm —— 本条只登记，未改原文件）'
  }
  const annotation = {
    annotation_id: createHash('sha256')
      .update(args.fingerprint + '|' + (args.note || '') + '|' + Date.now())
      .digest('hex').slice(0, 12),
    fingerprint: fpAtAnnotation,
    source_uri: e?.source_uri || args.source_uri || null,
    title: e ? (e.source_uri.split('/').pop() || '') : null,
    modality: e?.pool === 'not-material' ? 'unknown' : (e?.pool || 'unknown'),
    pool: e?.pool || null,
    format: e?.format || null,
    evidence_type: '原文',
    location: null,
    taxonomy_id: args.taxonomy_id || null,
    note: args.note || '',
    tags: args.tags || [],
    sensitivity: args.sensitivity || 'private',   // ★7：默认 private——缺省禁止云回退，放行需显式打 public
    created_at: new Date().toISOString(),
    author: args.author || 'agent',
    origin: 'source',          // evidence.py 语义：标注来自真实资料
    extraction_method: 'native',
    unverifiable,
    write_back: fileModified,
  }
  mkdirSync(dir, { recursive: true })
  const annPath = join(dir, 'annotations.jsonl')
  const existing = existsSync(annPath) ? readFileSync(annPath, 'utf-8') : ''
  writeFileSync(annPath, existing + JSON.stringify(annotation) + '\n', 'utf-8')
  return {
    ok: true,
    output: {
      root: rootReal, annotation, annotations_path: annPath,
      write_back: fileModified, wb_warning: wbWarning, unverifiable,
    },
  }
}

export async function kbVerifyImpl(args, helpers = {}) {
  let rootReal
  try {
    rootReal = realpathSync(args.root)
  } catch {
    return { ok: false, error: '[OPEN_FAILED] 目录不存在: ' + args.root }
  }
  const dir = join(STATE_DIR, 'kb', rootIdOf(rootReal))
  const problems = []
  const annotations = []
  const annPath = join(dir, 'annotations.jsonl')
  if (existsSync(annPath)) {
    readFileSync(annPath, 'utf-8').split(/\r?\n/).forEach((line, i) => {
      if (!line.trim()) return
      try {
        annotations.push(JSON.parse(line))
      } catch {
        problems.push({ kind: 'annotation-unparseable', line: i + 1 })
      }
    })
  }
  let registry = null
  try {
    registry = JSON.parse(readFileSync(join(dir, 'registry.json'), 'utf-8'))
  } catch {
    problems.push({ kind: 'registry-missing' })
  }
  // W2（工单 2026-10-06）：office_properties:true 时逐个只读打开 Office 标注文件，
  // 读回 kb_* 自定义属性与登记比对。比较基准=标注里**非空**的字段（标注没给的字段
  // 不要求文件里存在）；比对是只读动作，绝不写。
  const officeCheck = args.office_properties === true && typeof helpers.officeReader === 'function'
  let officeVerified = 0
  let officeChecked = 0
  // ★5 补（CC 2026-10-07，T4 实测暴露）：annotations.jsonl 是**追加式** —— 重标注会新增一行，
  // 而旧行仍在。逐行检查会让**任何写回过的文件永久报 fingerprint-changed**（重标注也救不回来，
  // 因为旧行一直在）⇒ 上轮"写回后重新标注即可"的说明是错的。
  // 判据改为：**每个源文件只看最新一行**。历史行是"当时那版文件"的历史记录，不代表现状；
  // 被取代的行数如实回报在 superseded。
  const latestBySrc = new Map()
  for (const a of annotations) {
    latestBySrc.set(a.source_uri || a.annotation_id, a)   // 后写覆盖先写 = 留最后一行
  }
  const superseded = annotations.length - latestBySrc.size
  let verified = 0
  for (const a of latestBySrc.values()) {
    const relPosix = a.source_uri?.split('://')[1]
    if (!relPosix) {
      problems.push({ kind: 'no-location', annotation_id: a.annotation_id })
      continue
    }
    const abs = join(rootReal, relPosix)
    let fpOk = false
    try {
      const s = statNs(abs)
      const fp = fingerprintOf(relPosix, s.size, s.mtimeNs)
      if (fp !== a.fingerprint) {
        problems.push({ kind: 'fingerprint-changed', annotation_id: a.annotation_id, source_uri: a.source_uri })
      } else {
        verified += 1
        fpOk = true
      }
    } catch {
      problems.push({ kind: 'source-gone', annotation_id: a.annotation_id, source_uri: a.source_uri })
      continue
    }
    if (!officeCheck) continue
    const ext = relPosix.toLowerCase().slice(relPosix.lastIndexOf('.'))
    if (!OFFICE_EXTS.includes(ext)) continue
    officeChecked += 1
    try {
      const r = await helpers.officeReader(abs)
      if (!r || r.error) {
        problems.push({ kind: 'office-read-failed', annotation_id: a.annotation_id, source_uri: a.source_uri, error: String(r?.error || 'reader failed').slice(0, 160) })
        continue
      }
      const props = r.properties || {}
      const expected = { kb_note: a.note, kb_taxonomy: a.taxonomy_id, kb_sensitivity: a.sensitivity }
      for (const [name, want] of Object.entries(expected)) {
        if (want === null || want === undefined || want === '') continue   // 标注没给的字段不要求
        if (!(name in props)) {
          problems.push({ kind: 'property-missing', annotation_id: a.annotation_id, name, source_uri: a.source_uri })
        } else if (String(props[name]) !== String(want)) {
          problems.push({ kind: 'property-mismatch', annotation_id: a.annotation_id, name,
                          expected: String(want), actual: String(props[name]).slice(0, 120) })
        }
      }
      officeVerified += 1
    } catch (e) {
      problems.push({ kind: 'office-read-failed', annotation_id: a.annotation_id, source_uri: a.source_uri, error: String(e).slice(0, 160) })
    }
  }
  let schemeConsistent = null
  for (const f of (existsSync(dir) ? readdirSync(dir) : [])) {
    if (!f.startsWith('scheme-') || !f.endsWith('.json')) continue
    try {
      const scheme = JSON.parse(readFileSync(join(dir, f), 'utf-8'))
      const missing = (scheme.moves || []).filter((m) => {
        const e = (registry?.entries || {})[m.fingerprint]
        const rel = m.from?.split('://')[1]
        return !e || !existsSync(join(rootReal, rel || ''))
      })
      if (missing.length) {
        problems.push({ kind: 'scheme-stale', scheme: f, missing: missing.length })
        schemeConsistent = false
      }
    } catch {
      problems.push({ kind: 'scheme-unreadable', scheme: f })
    }
  }
  return {
    ok: true,
    output: {
      root: rootReal, annotations: annotations.length, verified, problems, superseded,
      clean: problems.length === 0, scheme_consistent: schemeConsistent,
      ...(officeCheck ? { office_checked: officeChecked, office_verified: officeVerified } : {}),
    },
  }
}
