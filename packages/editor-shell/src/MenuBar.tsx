// 功能：菜单命令使用新核心已有事件入口和能力，默认切换后不再错误禁用；时间：2026-11-11；作者：lq。
/**
 * 经典菜单栏。
 *
 * 只在主题插件声明 `chrome.toolbar: 'classic-two-row'` 时渲染，
 * 用来还原 Office 2003 / XP 时期「标题栏 + 菜单栏 + 工具栏」的三段式结构。
 *
 * 设计约束：这里只放**应用级**命令。插入表格、字体、加粗这类文档级操作
 * 属于各编辑器自己的工具栏，不在外壳重复实现——否则两套入口会各自漂移。
 */
import { useMemo } from 'react';
import {
  BookOpen,
  Command,
  Eye,
  FilePlus2,
  FolderOpen,
  Info,
  Keyboard,
  Monitor,
  Moon,
  Package,
  Printer,
  Ruler,
  Save,
  Search,
  Settings,
  Sparkles,
  Sun,
  Wrench,
} from 'lucide-react';
import { useWorkspaceStore } from '@mini-office/document-model';
import { applyAppearance } from '@mini-office/design-system';
import { MOD_KEY, type OfficeKind } from '@mini-office/shared';
import { Dropdown, type MenuItem } from '@mini-office/ui';

const OFFICE_KINDS: Array<{ kind: OfficeKind; label: string }> = [
  { kind: 'docx', label: 'Word 文档' },
  { kind: 'xlsx', label: 'Excel 工作簿' },
  { kind: 'pptx', label: 'PowerPoint 演示文稿' },
];

function fire(name: string): void {
  window.dispatchEvent(new CustomEvent(name));
}

export function MenuBar({ onHome }: { onHome: () => void }) {
  const sessions = useWorkspaceStore((state) => state.sessions);
  const activeId = useWorkspaceStore((state) => state.activeId);
  const settings = useWorkspaceStore((state) => state.settings);
  const createDocument = useWorkspaceStore((state) => state.createDocument);
  const openFileDialog = useWorkspaceStore((state) => state.openFileDialog);
  const saveActive = useWorkspaceStore((state) => state.saveActive);
  const openDialog = useWorkspaceStore((state) => state.openDialog);
  const setCommandPalette = useWorkspaceStore((state) => state.setCommandPalette);
  const updateSettings = useWorkspaceStore((state) => state.updateSettings);

  const activeSession = activeId ? sessions[activeId] : undefined;
  const hasDocument = Boolean(activeSession);

  const menus = useMemo<Array<{ id: string; label: string; items: MenuItem[] }>>(() => {
    const setThemeMode = (theme: 'light' | 'dark' | 'system') => {
      applyAppearance({ theme, accent: settings.accent, customAccent: settings.customAccent, eyeCare: settings.eyeCare });
      updateSettings({ theme });
    };

    const create = async (kind: OfficeKind) => {
      await createDocument(kind);
      onHome();
    };

    return [
      {
        id: 'file',
        label: '文件',
        items: [
          { type: 'label', label: '新建' },
          ...OFFICE_KINDS.map(({ kind, label }) => ({
            id: `new-${kind}`,
            label: `新建${label}`,
            icon: <FilePlus2 size={13} />,
            onSelect: () => void create(kind),
          })),
          { type: 'separator' },
          {
            id: 'open',
            label: '打开…',
            icon: <FolderOpen size={13} />,
            shortcut: `${MOD_KEY}O`,
            onSelect: () => void openFileDialog(),
          },
          {
            id: 'save',
            label: '保存',
            icon: <Save size={13} />,
            shortcut: `${MOD_KEY}S`,
            disabled: !hasDocument,
            onSelect: () => void saveActive(),
          },
          { type: 'separator' },
          {
            id: 'export-pdf',
            label: '导出为 PDF',
            disabled: !hasDocument,
            onSelect: () => fire('mo:export-pdf'),
          },
          {
            id: 'print',
            label: '打印',
            icon: <Printer size={13} />,
            shortcut: `${MOD_KEY}P`,
            disabled: !hasDocument,
            onSelect: () => fire('mo:print'),
          },
          { type: 'separator' },
          {
            id: 'home',
            label: '返回首页',
            icon: <BookOpen size={13} />,
            onSelect: onHome,
          },
        ],
      },
      {
        id: 'edit',
        label: '编辑',
        items: [
          {
            id: 'find',
            label: '查找与替换',
            icon: <Search size={13} />,
            shortcut: `${MOD_KEY}F`,
            disabled: !hasDocument,
            onSelect: () => fire('mo:find'),
          },
          {
            // 功能：经典外观下也能直接打开 Word 样式库；时间：2026-10-29；作者：lq。
            id: 'styles',
            label: '样式库…',
            icon: <Sparkles size={13} />,
            disabled: !hasDocument,
            onSelect: () => fire('mo:styles'),
          },
          { type: 'separator' },
          {
            id: 'palette',
            label: '命令面板',
            icon: <Command size={13} />,
            shortcut: `${MOD_KEY}K`,
            onSelect: () => setCommandPalette(true),
          },
        ],
      },
      {
        id: 'view',
        label: '视图',
        items: [
          {
            id: 'theme-light',
            label: '浅色',
            icon: <Sun size={13} />,
            checked: settings.theme === 'light',
            onSelect: () => setThemeMode('light'),
          },
          {
            id: 'theme-dark',
            label: '深色',
            icon: <Moon size={13} />,
            checked: settings.theme === 'dark',
            onSelect: () => setThemeMode('dark'),
          },
          {
            id: 'theme-system',
            label: '跟随系统',
            icon: <Monitor size={13} />,
            checked: settings.theme === 'system',
            onSelect: () => setThemeMode('system'),
          },
          { type: 'separator' },
          {
            id: 'toggle-toolbar',
            label: '显示工具栏',
            checked: settings.showToolbar,
            onSelect: () => updateSettings({ showToolbar: !settings.showToolbar }),
          },
          {
            id: 'toggle-properties',
            label: '属性面板',
            checked: settings.showProperties,
            onSelect: () => updateSettings({ showProperties: !settings.showProperties }),
          },
          // 功能：标尺默认隐藏，与 Word / WPS 的「视图 → 标尺」一致，需要精确缩进与制表位时再开；时间：2026-10-30；作者：lq。
          {
            id: 'toggle-ruler',
            label: '标尺',
            icon: <Ruler size={13} />,
            checked: settings.wordRuler === true,
            onSelect: () => updateSettings({ wordRuler: settings.wordRuler !== true }),
          },
          {
            // 功能：裁剪标记独立于编辑标记和标尺，按 Word 习惯单独控制；时间：2026-11-05；作者：lq。
            id: 'toggle-word-crop-marks',
            label: '裁剪标记',
            checked: settings.wordCropMarks !== false,
            onSelect: () => updateSettings({ wordCropMarks: settings.wordCropMarks === false }),
          },
          {
            id: 'toggle-eyecare',
            label: '护眼模式',
            icon: <Eye size={13} />,
            checked: settings.eyeCare,
            onSelect: () => {
              const eyeCare = !settings.eyeCare;
              applyAppearance({
                theme: settings.theme,
                accent: settings.accent,
                customAccent: settings.customAccent,
                eyeCare,
              });
              updateSettings({ eyeCare });
            },
          },
        ],
      },
      {
        id: 'tools',
        label: '工具',
        items: [
          {
            id: 'plugins',
            label: '插件…',
            icon: <Package size={13} />,
            onSelect: () => openDialog({ type: 'plugins' }),
          },
          {
            id: 'settings',
            label: '设置…',
            icon: <Settings size={13} />,
            shortcut: `${MOD_KEY},`,
            onSelect: () => openDialog({ type: 'settings' }),
          },
        ],
      },
      {
        id: 'help',
        label: '帮助',
        items: [
          {
            id: 'shortcuts',
            label: '键盘快捷键',
            icon: <Keyboard size={13} />,
            onSelect: () => openDialog({ type: 'shortcuts' }),
          },
          {
            id: 'compatibility',
            label: '兼容性报告',
            icon: <Wrench size={13} />,
            disabled: !hasDocument,
            onSelect: () => {
              if (activeId) openDialog({ type: 'compatibility', sessionId: activeId });
            },
          },
          { type: 'separator' },
          {
            id: 'about',
            label: '关于 Mini Office',
            icon: <Info size={13} />,
            onSelect: () => openDialog({ type: 'about' }),
          },
        ],
      },
    ];
  }, [
    activeId,
    createDocument,
    hasDocument,
    onHome,
    openDialog,
    openFileDialog,
    saveActive,
    setCommandPalette,
    settings.accent,
    settings.customAccent,
    settings.eyeCare,
    settings.showProperties,
    settings.showToolbar,
    settings.theme,
    settings.wordRuler,
    settings.wordCropMarks,
    updateSettings,
  ]);

  return (
    <nav aria-label="菜单栏" className="mo-menubar shrink-0">
      {menus.map((menu) => (
        <Dropdown
          key={menu.id}
          items={menu.items}
          minWidth={216}
          trigger={({ open, toggle }) => (
            <button
              type="button"
              onClick={toggle}
              aria-haspopup="menu"
              aria-expanded={open}
              data-active={open ? 'true' : undefined}
              className="mo-menubar-item"
            >
              {menu.label}
            </button>
          )}
        />
      ))}
    </nav>
  );
}
