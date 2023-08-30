using System.Runtime.InteropServices;
using System.Text;
using Spectre.Console;
using Spectre.Console.Rendering;

namespace AlwaysOnTop;

internal static class Program
{
    #region DLL Imports
    #pragma warning disable SYSLIB1054 // Use 'LibraryImportAttribute' instead of 'DllImportAttribute' to generate P/Invoke marshalling code at compile time
    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SetWindowPos(nint hWnd, nint hWndInsertAfter, int x, int y, int cx, int cy,
        uint uFlags);
    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(nint hwndParent, EnumWindowsProc lpEnumFunc, nint lParam);
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, nint lParam);
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    private static extern int GetWindowTextLength(nint hWnd);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(nint hWnd, StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern int GetClassName(nint hWnd, StringBuilder lpClassName, int nMaxCount);
    #pragma warning restore SYSLIB1054 // Use 'LibraryImportAttribute' instead of 'DllImportAttribute' to generate P/Invoke marshalling code at compile time
    #endregion


    private delegate bool EnumWindowsProc(nint hWnd, nint lParam);

    private static readonly List<WindowInfo> WindowData = new();
    private record WindowInfo(nint Handle, string Title, string ClassName);

    private static readonly nint HwndTopmost = new(-1);
    private static readonly nint HwndNotopmost = new(-2);
    private const uint SwpNosize = 0x0001;
    private const uint SwpNomove = 0x0002;
    private const string PromptHandleMessage = "Enter the [green]handle[/]";
    private const string HandleParseFailureMessage = "[red]Failed to parse handle.[/]";
    private const string ValidationError = "Please enter a valid handle";
    private const string AppName = "AlwaysOnTop";
    private const string Green = "green";
    private const int DefaultStringBuilderCapacity = 256;
    private const int DisplayRefreshDelay = 1000;



    private static void Main() => MainAsync().GetAwaiter().GetResult();

    private static async Task MainAsync()
    {
        DisplayHeader();

        while (true)
        {
            PopulateWindowData();
            await DisplayWindows();
            await SetSelectedWindowAlwaysOnTop();
        }
    }

    private static Task SetSelectedWindowAlwaysOnTop()
    {
        var handleInput = PromptForHandle();
        if (!TryParseHandle(handleInput, out var selectedHandle))
        {
            DisplayHandleParseError();
            return Task.CompletedTask;
        }

        var choice = PromptForAlwaysOnTopConfirmation(selectedHandle);
        SetWindowToAlwaysOnTop(selectedHandle, choice);
        return Task.CompletedTask;
    }

    private static string PromptForHandle()
    {
        return AnsiConsole.Prompt(
            new TextPrompt<string>(PromptHandleMessage)
                .PromptStyle(Green)
                .Validate(input => nint.TryParse(input, out _) ? ValidationResult.Success() : ValidationResult.Error(ValidationError)));
    }

    private static bool TryParseHandle(string handleInput, out nint handle)
    {
        return nint.TryParse(handleInput, out handle);
    }

    private static void DisplayHandleParseError()
    {
        AnsiConsole.WriteLine(HandleParseFailureMessage);
    }

    private static void DisplayHeader()
    {
        AnsiConsole.Write(new FigletText(AppName).Centered().Color(Color.Green));
        AnsiConsole.WriteLine();
    }

    private static nint PromptForAlwaysOnTopConfirmation(nint handle)
    {
        var choice = AnsiConsole.Prompt(
            new SelectionPrompt<string>()
                .Title($"Set {handle} to Always be on top?")
                .PageSize(3)
                .AddChoices("Yes", "No"));

        var alwaysOnTop = choice == "Yes" ? HwndTopmost : HwndNotopmost;
        return alwaysOnTop;
    }

    private static void SetWindowToAlwaysOnTop(nint handle, nint choice)
    {
        SetWindowPos(handle, choice, 0, 0, 0, 0, SwpNomove | SwpNosize);
    }
    private static async Task DisplayWindows()
    {
        var table = CreateWindowDataTable();
        await DisplayTableWithLiveRefresh(table);
    }

    private static Table CreateWindowDataTable()
    {
        var table = new Table().Centered().Border(TableBorder.Rounded);
        table.AddColumn("Handle");
        table.AddColumn("Title");
        table.AddColumn("Class");

        foreach (var window in WindowData)
        {
            table.AddRow(window.Handle.ToString(), Markup.Escape(window.Title), Markup.Escape(window.ClassName));
        }

        return table;
    }

    private static async Task DisplayTableWithLiveRefresh(IRenderable table)
    {
        await AnsiConsole.Live(table)
            .AutoClear(false)
            .Overflow(VerticalOverflow.Ellipsis)
            .Cropping(VerticalOverflowCropping.Top)
            .StartAsync(async ctx =>
            {
                ctx.Refresh();
                await Task.Delay(DisplayRefreshDelay);
            });
    }

    private static void PopulateWindowData()
    {
        WindowData.Clear();
        EnumerateWindowsAndPopulateData(nint.Zero);
    }

    private static void EnumerateWindowsAndPopulateData(nint parentHwnd)
    {
        EnumWindows((hWnd, lParam) =>
        {
            AddWindowDataIfValid(hWnd);
            EnumChildWindows(hWnd, (childHWnd, childParam) =>
            {
                AddWindowDataIfValid(childHWnd);
                return true;
            }, nint.Zero);
            return true;
        }, parentHwnd);
    }

    private static void AddWindowDataIfValid(nint hWnd)
    {
        var windowInfo = GetWindowInfo(hWnd);
        if (!string.IsNullOrEmpty(windowInfo.Title) || !string.IsNullOrEmpty(windowInfo.ClassName))
        {
            WindowData.Add(windowInfo);
        }
    }

    private static WindowInfo GetWindowInfo(nint hWnd)
    {
        var title = GetWindowTitle(hWnd);
        var className = GetWindowClassName(hWnd);
        return new WindowInfo(hWnd, title, className);
    }

    private static string GetWindowTitle(nint hWnd)
    {
        var length = GetWindowTextLength(hWnd);
        if (length <= 0)
        {
            return string.Empty;
        }
        var titleBuilder = new StringBuilder(length + 1);
        var result = GetWindowText(hWnd, titleBuilder, titleBuilder.Capacity);
        return result == 0 ? string.Empty : titleBuilder.ToString();
    }

    private static string GetWindowClassName(nint hWnd)
    {
        var classBuilder = new StringBuilder(DefaultStringBuilderCapacity);
        var result = GetClassName(hWnd, classBuilder, classBuilder.Capacity);
        return result == 0 ? string.Empty : classBuilder.ToString();
    }
}