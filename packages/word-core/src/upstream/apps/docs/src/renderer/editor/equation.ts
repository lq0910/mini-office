/** 功能：融合 GenOffice Word 内核；时间：2026-11-08；集成作者：lq；上游作者、固定版本及本地修改见 packages/word-core/UPSTREAM.md 和 SOURCE-MANIFEST.json。 */
import { latexToOmml, mathParagraphXml, mathTokensOf, ommlToMathML } from '../../../../../packages/docx-engine/src/index'
import { t } from '../i18n/locale'
import type { PmNode } from './convert'

/**
 * Protected display-equation block built from LaTeX (throws on syntax outside
 * the supported subset). Shared by the insert dialog/gallery and the AI
 * <formula> tag; genXml carries the OMML through the save path.
 */
/** MathML for inline flow (ommlToMathML emits display="block" for equations);
 *  `display` keeps display math-style (large operators, stacked limits) in the flow */
export function inlineMathML(omml: string, display = false): string {
  return ommlToMathML(omml).replace(
    / display="block"/g,
    display ? ' display="inline" displaystyle="true"' : ' display="inline"',
  )
}

/**
 * Atomic inline formula node built from LaTeX (throws on unsupported syntax).
 * Flows with paragraph text; saves as a Run.math (<m:oMath> emitted verbatim).
 */
export function inlineEquationNodeJson(latex: string): PmNode {
  const inner = latexToOmml(latex)
  const omml = `<m:oMath>${inner}</m:oMath>`
  return {
    type: 'docInlineMath',
    attrs: {
      omml,
      mathml: inlineMathML(omml),
      latex: latex.trim(),
      text: mathTokensOf(omml).join(''),
    },
  }
}

export function equationBlockJson(latex: string): PmNode {
  const omml = latexToOmml(latex)
  return {
    type: 'docProtected',
    attrs: {
      docxIndex: null,
      blockType: 'passthrough',
      label: t('editorEquation'),
      previewText: latex.trim(),
      genXml: mathParagraphXml(omml),
      formulaDisplay: {
        tokens: mathTokensOf(omml),
        mathml: ommlToMathML(`<m:oMath>${omml}</m:oMath>`),
        omml: `<m:oMath>${omml}</m:oMath>`,
        latex: latex.trim(),
      },
    },
  }
}
