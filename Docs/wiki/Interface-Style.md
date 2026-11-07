# Interface style

The user's direction is to follow the launcher's original palette exactly. The source of truth is the pre-upgrade `$Global:Themes.Dracula` block, verified against both the saved script and Git's original source. The theme name is not a reason to substitute a different palette. Layout and keyboard-focus improvements retain these existing colors.

## Dracula tokens

| Role | Color |
| --- | --- |
| Window / title bar | `#282A36` |
| File list, fields, activity surface | `#202020` |
| Buttons | `#2F3440` |
| Border / selected row | `#44475A` |
| Existing slate-blue accent / keyboard focus | `#6272A4` |
| Main text | `#F8F8F2` |
| Secondary text | `#969696` |

Buttons retain their original dark fill, off-white text, and slate-blue outlines. No violet fill or hover color is used. Light mode retains its original separate palette.

## Layout and controls

- Keep the result list and activity panel prominent. Use the folder name as the heading and a quieter full path below it.
- Paint the complete native header, including its empty trailing area. Keep column dividers and the list/activity resize divider visible in the existing border color. Recover collapsed saved columns and fit them inside the list. Full-path tooltips preserve access to truncated paths.
- Use a flat, quiet field surface with a visible colored keyboard-focus border.
- Keep Run first in the action row. Its caption reflects administrator mode; its colors follow the original button style.
- Keep launch options, Settings, Open Log, and Copy Error under More. F2, the context menu, and the visible Shortcut Name button still edit only a future label.
- Edit shortcut names in a themed field aligned with the existing text after its icon. Hide the original label while editing; keep the field inside its column when resized. Enter saves, Escape cancels, and focus loss saves.
- Retain native keyboard behavior, resizable windows, DPI configuration, and the actual file identity column.
- Avoid decorative animation, bright frames around every panel, and a long row of equally emphasized utility buttons.
- Cleanup uses three independent checkboxes for source files, extracted files, and the qBittorrent entry. Show each location once, followed by a short effect summary. Acknowledgement appears only when the extracted app/installer will be removed. There is no queue window.

`Tests/Preview-Interface.ps1` captures representative real forms in either theme using disposable files. `Tests/Test-UI.ps1` checks compact options remain reachable, controls fit the layout, aliases preserve original files, and cleanup acknowledgement works.
