import { HardBreak } from '@tiptap/extension-hard-break'
import { mergeAttributes } from '@tiptap/core'

export const CustomHardBreak = HardBreak.extend({
  addKeyboardShortcuts() {
    return {
      'Shift-Enter': () => this.editor.commands.setHardBreak(),
      'Mod-Enter': () => this.editor.commands.setHardBreak(),
    }
  },

  addNodeView() {
    return () => {
      return {
        dom: document.createElement('br'),
      }
    }
  },

  addInputRules() {
    // Prevent automatic hard break insertion on typing
    return []
  },

  parseHTML() {
    return [
      {
        tag: 'br',
        getAttrs: element => {
          const parent = element.parentElement
          if (parent && parent.tagName === 'P' && parent.childNodes.length === 1) {
            // Prevent <br> from being parsed when it's the only child in a <p>
            return false
          }
          return {}
        },
      },
    ]
  },

  renderHTML({ HTMLAttributes }) {
    return ['br', mergeAttributes(HTMLAttributes)]
  },
})

