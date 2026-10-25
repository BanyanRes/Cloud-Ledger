# Herdr cheatsheet

Herdr is mouse-first: click panes/tabs to focus, drag borders to resize, right-click for menus. Keys below are optional.

## Prefix keys (press Ctrl+B, let go, then the key)

| Keys | What it does |
|---|---|
| Ctrl+B then V | split pane to the right |
| Ctrl+B then - | split pane down |
| Ctrl+B then Z | zoom / unzoom the current pane |
| Ctrl+B then C | new tab |
| Ctrl+B then Q | detach — closes the view, agents keep running |
| Ctrl+B then ? | show every key binding |

Korean keyboard tip: switch to English input (한/영) before the prefix key, or the letter after Ctrl+B may not register. Herdr has an experimental fix you can turn on in `%APPDATA%\herdr\config.toml`:

```toml
[experimental]
switch_ascii_input_source_in_prefix = true
```

## Paste and copy (Windows Terminal)

- Copy: drag-select text inside a pane.
- Paste: Ctrl+Shift+V.

## Get back in

- `herdr` — reattach (run it from a normal terminal, not inside a Herdr pane).

## Agent status in the sidebar

| State | Meaning | What to do |
|---|---|---|
| working | busy | wait |
| blocked | waiting for YOUR answer (approval/question) | look at that pane |
| done | finished, not looked at yet | review it |
| idle | finished / ready for input | nothing |
| unknown | Herdr can't tell | look at the pane |

The dot colors come from the theme. If they're hard to tell apart, switch to distinct symbols in `%APPDATA%\herdr\config.toml`, then run `herdr server reload-config`:

```toml
[ui]
status_indicators = "symbols"
```

From a terminal, `herdr agent list` shows every agent and its state.

## Fully stop everything

`herdr server stop` — stops the background server AND every pane/agent in it. Detaching (Ctrl+B then Q) does not stop anything.