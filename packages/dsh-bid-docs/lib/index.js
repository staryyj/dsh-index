/**
 * 随包分发的 bid-docs 技能：投标技术标 + 附表生成。
 * @module dsh-bid-docs
 */
import { readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'

const PROVIDER_NAME = 'bid-docs'
const SKILL_BODY_URL = new URL('../assets/SKILL.md', import.meta.url)
const RESOURCE_BASE = { kind: 'directory', path: fileURLToPath(new URL('../assets/', import.meta.url)) }
const INVOCATION = { modelInvocable: true, userInvocable: true }
const DESCRIPTION = '从招标资料压缩包自动生成 A3 四页技术标（施工组织设计）与附表（机械设备/劳动力/进度横道图/总平面+临时用地），含招标参数抽取与清单工程量核算。当用户发来招标文件包并要求出技术标、施工组织设计、A3 方案或附表时使用。'
const CANDIDATE = {
  name: 'bid-docs',
  description: DESCRIPTION,
  invocation: INVOCATION,
  provider: PROVIDER_NAME,
  source: 'bundled',
  resourceBase: RESOURCE_BASE,
  rank: 600,
  locator: SKILL_BODY_URL,
}

const provider = {
  name: PROVIDER_NAME,
  list: () => Promise.resolve([CANDIDATE]),
  async get() {
    return {
      name: CANDIDATE.name,
      description: CANDIDATE.description,
      invocation: CANDIDATE.invocation,
      provider: CANDIDATE.provider,
      source: CANDIDATE.source,
      resourceBase: RESOURCE_BASE,
      content: await readFile(SKILL_BODY_URL, 'utf8'),
    }
  },
}

export const name = 'bid-docs'
export const inject = ['skills']
export function apply(ctx) {
  ctx.skills.registerProvider(() => provider)
}
