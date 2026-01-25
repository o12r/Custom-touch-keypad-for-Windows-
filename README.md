# Overview: 
This project is a lightweight, fully customizable on-screen keypad designed for Windows. It allows users to create personalized grids of buttons that send keystrokes to any active application.

- Non-Intrusive Overlay: Stays on top of other windows without stealing focus, ensuring your keystrokes go exactly where you need them.

- Visual Design Mode: Drag-and-drop resizing, intuitive color themes, and easy key assignment.

- Slide Mode: Choose between standard typing or slide mode for rapid inputs.

- Config Profiles: Save and load layouts for different games or software.

# User Instructions
## Installation & Requirements
You will need Python installed. Before running the script, install the required dependencies via your terminal/command prompt:

`
Bash
`
`
pip install keyboard pywin32
`


## Design Mode (Edit)
When you launch the app, it starts in Design Mode. This is where you build your layout.

Assigning Keys: Click any cell in the grid. A text box will appear. Type the key or shortcut you want (e.g., a, ctrl+c, F1) and press Enter or click the next cell to edit.

Resizing Rows/Columns:

Drag: Click and drag the gray number bars (headers) on the top or left to resize rows and columns visually.

Precise Edit: Click a number header once to type an exact size value.

Adding/Removing Cells: Use the Control Panel buttons (+ Row, - Col, etc.) to expand or shrink your grid.

Merging Cells: Hold Right-Click and drag across multiple cells to select them. Click Merge in the panel to combine them into one large button.

## Lock Mode (Play)
To use your controller, click the Lock button on the Control Panel.

What happens: The title bar and grid headers disappear. The window becomes a floating overlay.

How to use: Click the buttons to send keystrokes to your active window (e.g., Notepad, a Game, Browser). The app will not steal focus, so your game/doc remains active.

To Edit Again: A small red Unlock icon will appear (usually floating near your grid). Drag it slightly to unlock the grid and return to Design Mode.

## Advanced Settings
Rapid Fire (Slide Mode):

Type : Standard button behavior. Click to type once.

Slide : Drag over buttons to fire rapidly and continuously as long as you hover over them. Great for rhythm games or spamming inputs.

Opacity: Use the slider in the panel to make the window semi-transparent.

Themes: Click Theme  to switch between Light, Dark, or System themes, or use Color to pick a custom accent color.
