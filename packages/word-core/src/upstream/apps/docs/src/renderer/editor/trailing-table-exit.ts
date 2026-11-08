/** 功能：融合 GenOffice Word 内核；时间：2026-11-08；集成作者：lq；上游作者、固定版本及本地修改见 packages/word-core/UPSTREAM.md 和 SOURCE-MANIFEST.json。 */
import { Extension } from '@tiptap/core'
import { GapCursor } from '@tiptap/pm/gapcursor'
import { Plugin, PluginKey, TextSelection } from '@tiptap/pm/state'
import { TRACK_IGNORE } from './revisions'

const key = new PluginKey('trailingTableExit')

export const TrailingTableExitExtension = Extension.create({
  name: 'trailingTableExit',

  addProseMirrorPlugins() {
    const editor = this.editor
    return [
      new Plugin({
        key,
        appendTransaction(transactions, oldState, newState) {
          if (
            !editor.view.editable ||
            oldState.doc !== newState.doc ||
            !(newState.selection instanceof GapCursor) ||
            newState.selection.from !== newState.doc.content.size ||
            newState.doc.lastChild?.type.name !== 'docTable' ||
            !transactions.some((transaction) => transaction.selectionSet && !transaction.docChanged)
          ) {
            return null
          }

          const paragraph = newState.schema.nodes.docParagraph?.createAndFill()
          if (!paragraph) return null
          const pos = newState.doc.content.size
          const transaction = newState.tr.insert(pos, paragraph)
          // leaving the table is navigation, not authoring: no revision mark
          return transaction
            .setMeta(TRACK_IGNORE, true)
            .setSelection(TextSelection.create(transaction.doc, pos + 1))
            .scrollIntoView()
        },
      }),
    ]
  },
})
