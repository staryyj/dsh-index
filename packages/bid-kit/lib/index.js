/**
 * bid-kit：投标全流程工具包（多技能随包分发）
 * @module bid-kit
 */
import { readFile, readdir } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import { basename, join } from 'node:path'

const PROVIDER_NAME = 'bid-kit'
const ASSETS = fileURLToPath(new URL('../assets/', import.meta.url))
const SKILL_DIR = join(ASSETS, 'skills')
const RESOURCE_BASE = { kind: 'directory', path: ASSETS }
const INVOCATION = { modelInvocable: true, userInvocable: true }
const RANK = 600

const TITLES = {
  'bid-docs': '从招标资料压缩包自动生成 A3 四页技术标（施工组织设计）与附表（机械设备/劳动力/进度横道图/总平面+临时用地），含招标参数抽取与清单工程量核算',
  'bid-tender': '解析招标文件：抽取工期/造价/限价/保证金/资质/建造师/技术标页数/暗标要求，识别废标红线与否决条款，并与旧投标模板逐页对照',
  'bid-package': '投标文件整册组装：章节顺序、封面目录页码、暗标去痕、多标段交叉检查、最终 PDF 交付与命名规范',
  'bid-pricing': '报价与资信核查：最高投标限价、安全文明施工费下限、暂列金额、造价工程师签章、投标有效期、保证金、业绩与财务要求',
  'doc-render': '文档渲染与页面校核：docx 转 PDF、统计页数、逐页截图核对、图片尺寸与页边距检查、LibreOffice Kit 用法速查',
}

let candidates = null

async function loadCandidates() {
  if (candidates) return candidates
  let files = []
  try {
    files = (await readdir(SKILL_DIR)).filter((f) => f.endsWith('.md'))
  } catch {
    files = []
  }
  candidates = files.map((file) => {
    const name = basename(file, '.md')
    return {
      name,
      description: TITLES[name] || '投标工具',
      invocation: INVOCATION,
      provider: PROVIDER_NAME,
      source: 'bundled',
      resourceBase: RESOURCE_BASE,
      rank: RANK,
      locator: new URL(`../assets/skills/${file}`, import.meta.url),
    }
  })
  return candidates
}

const provider = {
  name: PROVIDER_NAME,
  async list() {
    return loadCandidates()
  },
  async get(candidate) {
    return {
      name: candidate.name,
      description: candidate.description,
      invocation: candidate.invocation,
      provider: candidate.provider,
      source: candidate.source,
      resourceBase: RESOURCE_BASE,
      content: await readFile(candidate.locator, 'utf8'),
    }
  },
}

export const name = 'bid-kit'
export const inject = ['skills']
export function apply(ctx) {
  ctx.skills.registerProvider(() => provider)
}
