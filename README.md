# .NET DarkMode

Adds support for Windows 10 dark mode in .NET applications. Based on https://github.com/ysc3839/win32-darkmode.

## Usage

Dark Mode can be enabled for your project in 2 easy steps:

### Step 1: Enabling dark mode for your application

1. Add a reference to the DarkMode.dll assembly in your application.
2. Go to your project properties, and click View Application Events.
3. Paste in the following code in the Application class:

- C#:

```csharp
private void Application_Startup(Object sender, StartupEventArgs e) {
    DarkMode.DarkMode.SetAppTheme(DarkMode.DarkMode.Theme.SYSTEM);
}
```

- Visual Basic:

```vb
Private Sub Application_Startup(sender As Object, e As StartupEventArgs) Handles MyBase.Startup
    DarkMode.DarkMode.SetAppTheme(DarkMode.DarkMode.Theme.SYSTEM)
End Sub
```

### Step 2: Enabling dark mode for each window

1. Open the class for one of your Forms.
2. Add the following code:

- C#:

```csharp
protected override void WndProc(ref Message m) {
    DarkMode.WndProc(this, m, DarkMode.DarkMode.Theme.SYSTEM);
    base.WndProc(m);
}
```

- Visual Basic:

```vb
Protected Overrides Sub WndProc(ByRef m As Message)
    DarkMode.WndProc(Me, m, DarkMode.DarkMode.Theme.SYSTEM)
    MyBase.WndProc(m)
End Sub
```

3. Repeat for every other Window you want to enable dark mode for.

## Themes

This library supports three themes:

|  Theme   | Description                            |
|:--------:|:---------------------------------------|
| `SYSTEM` | Respects the system dark mode setting. |
| `DARK`   | Forces dark mode.                      |
| `LIGHT`  | Forces light mode.                     |

## Other Methods

### `UpdateWindowTheme()`

Updates the theme for the specified window.

**Signature**:

- C#:

```csharp
public void UpdateWindowTheme(Form window, Theme theme)
```

- Visual Basic:

```vb
Public Sub UpdateWindowTheme(window As Form, theme As Theme)
```

**Parameters**:

- `window`: The window to update the theme for.
- `theme`: The new theme.
