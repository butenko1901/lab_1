using MyExcelMAUIApp.Parsing;

namespace MyExcelMAUIApp;

public partial class MainPage : ContentPage
{
    const int InitialRows = 50;
    const int InitialColumns = 20;
    const double CellWidth = 110, CellHeight = 40, HeaderWidth = 40, HeaderHeight = 30;

    const bool RequireLogical = false;

    Spreadsheet _sheet = NewSheet(InitialRows, InitialColumns);
    readonly Dictionary<(int row, int col), Entry> _entries = new();
    readonly List<Label> _rowLabels = new();
    readonly List<Label> _colLabels = new();
    bool _showValues;                       
    (int row, int col)? _selected;          

    public MainPage()
    {
        InitializeComponent();
        BuildGrid();
    }

    static Spreadsheet NewSheet(int rows, int cols) =>
        new Spreadsheet(rows, cols) { RequireLogicalResult = RequireLogical };


    void BuildGrid()
    {
        grid.Children.Clear();
        grid.RowDefinitions.Clear();
        grid.ColumnDefinitions.Clear();
        _entries.Clear();
        _rowLabels.Clear();
        _colLabels.Clear();

        // рядок 0 — підписи стовпчиків, стовпчик 0 — номери рядків
        grid.RowDefinitions.Add(new RowDefinition(new GridLength(HeaderHeight)));
        grid.ColumnDefinitions.Add(new ColumnDefinition(new GridLength(HeaderWidth)));

        for (int r = 1; r <= _sheet.RowCount; r++) AddRowHeader(r);
        for (int c = 1; c <= _sheet.ColumnCount; c++) AddColumnHeader(c);
        for (int r = 1; r <= _sheet.RowCount; r++)
            for (int c = 1; c <= _sheet.ColumnCount; c++)
                CreateEntry(r, c);
    }
   void AddRowHeader(int row)
{
    grid.RowDefinitions.Add(new RowDefinition(new GridLength(CellHeight)));
    var label = new Label
    {
        Text = row.ToString(),
        BackgroundColor = Color.FromArgb("#EBD3DA"),
        TextColor = Color.FromArgb("#6B5A60"),
        FontSize = 12,
        HorizontalTextAlignment = TextAlignment.Center,
        VerticalTextAlignment = TextAlignment.Center
    };
    Grid.SetRow(label, row);
    Grid.SetColumn(label, 0);
    grid.Children.Add(label);
    _rowLabels.Add(label);
}

void AddColumnHeader(int col)
{
    grid.ColumnDefinitions.Add(new ColumnDefinition(new GridLength(CellWidth)));
    var label = new Label
    {
        Text = Spreadsheet.ColumnName(col),
        BackgroundColor = Color.FromArgb("#EBD3DA"),
        TextColor = Color.FromArgb("#6B5A60"),
        FontSize = 12,
        HorizontalTextAlignment = TextAlignment.Center,
        VerticalTextAlignment = TextAlignment.Center
    };
    Grid.SetRow(label, 0);
    Grid.SetColumn(label, col);
    grid.Children.Add(label);
    _colLabels.Add(label);
}
    void CreateEntry(int row, int col)
    {
        var entry = new Entry
        {
            Text = "",
            IsReadOnly = _showValues,
            VerticalOptions = LayoutOptions.Center
        };
        entry.Focused += Entry_Focused;
        entry.Unfocused += Entry_Unfocused;
        Grid.SetRow(entry, row);
        Grid.SetColumn(entry, col);
        grid.Children.Add(entry);
        _entries[(row, col)] = entry;
    }

    void RemoveEntry(int row, int col)
    {
        if (_entries.Remove((row, col), out var entry))
            grid.Children.Remove(entry);
    }

    void CommitEntries()
    {
        if (_showValues) return;
        foreach (var (key, entry) in _entries)
            _sheet.SetExpression(key.row, key.col, entry.Text);
    }
    void RefreshDisplay()
    {
        _sheet.Recalculate();

        foreach (var (key, entry) in _entries)
        {
            entry.IsReadOnly = _showValues;

            if (_showValues)
            {
                var result = _sheet.Evaluate(key.row, key.col);
                entry.Text = result.Display;
                if (result.IsError) entry.TextColor = Colors.Red;
                else entry.ClearValue(Entry.TextColorProperty);
            }
            else
            {
                entry.Text = _sheet.GetExpression(key.row, key.col);
                entry.ClearValue(Entry.TextColorProperty);
            }
        }

        modeButton.Text = _showValues ? "Режим: ЗНАЧЕННЯ" : "Режим: ВИРАЗ";
        UpdateStatus();
    }

    void UpdateStatus()
    {
        if (!_selected.HasValue) { statusLabel.Text = ""; return; }

        var (r, c) = _selected.Value;
        string name = Spreadsheet.CellName(r, c);
        var result = _sheet.Evaluate(r, c);

        if (result.IsError)
        {
            statusLabel.TextColor = Colors.Red;
            statusLabel.Text = $"{name}: {result.Error} (позиція {result.ErrorPosition + 1})";
        }
        else
        {
            statusLabel.ClearValue(Label.TextColorProperty);
            statusLabel.Text = result.IsEmpty ? $"{name}: порожньо" : $"{name} = {result.Display}";
        }
    }


    void Entry_Focused(object? sender, FocusEventArgs e)
    {
        var entry = (Entry)sender!;
        _selected = (Grid.GetRow(entry), Grid.GetColumn(entry));
        textInput.Text = _sheet.GetExpression(_selected.Value.row, _selected.Value.col);
        UpdateStatus();
    }

    void Entry_Unfocused(object? sender, FocusEventArgs e)
    {
        var entry = (Entry)sender!;
        if (!_showValues)
            _sheet.SetExpression(Grid.GetRow(entry), Grid.GetColumn(entry), entry.Text);
        UpdateStatus();
    }


    void TextInput_Completed(object? sender, EventArgs e)
    {
        if (_showValues || !_selected.HasValue) return;

        var (r, c) = _selected.Value;
        _sheet.SetExpression(r, c, textInput.Text);
        _entries[(r, c)].Text = _sheet.GetExpression(r, c);
        UpdateStatus();
    }

    void CalculateButton_Clicked(object? sender, EventArgs e)
    {
        CommitEntries();
        _showValues = true;
        RefreshDisplay();
    }

    void ModeButton_Clicked(object? sender, EventArgs e)
    {
        CommitEntries();
        _showValues = !_showValues;
        RefreshDisplay();
    }


    async void SaveButton_Clicked(object? sender, EventArgs e)
    {
        CommitEntries();

        string? name = await DisplayPromptAsync("Збереження", "Введіть назву файлу:",
            "Зберегти", "Скасувати", initialValue: "table");
        if (string.IsNullOrWhiteSpace(name)) return;

        name = string.Concat(name.Trim().Split(Path.GetInvalidFileNameChars()));
        if (name.Length == 0) return;
        if (!name.EndsWith(".json", StringComparison.OrdinalIgnoreCase)) name += ".json";

        string path = Path.Combine(FileSystem.AppDataDirectory, name);
        try
        {
            await File.WriteAllTextAsync(path, _sheet.ToJson());
            await DisplayAlert("Збережено", $"Таблицю збережено у файл:\n{path}", "OK");
        }
        catch (Exception ex)
        {
            await DisplayAlert("Помилка", $"Не вдалося зберегти файл: {ex.Message}", "OK");
        }
    }

    async void ReadButton_Clicked(object? sender, EventArgs e)
    {
        string dir = FileSystem.AppDataDirectory;
        string[] files = Directory.GetFiles(dir, "*.json")
            .Select(f => Path.GetFileName(f) ?? f)
            .OrderBy(f => f)
            .ToArray();

        if (files.Length == 0)
        {
            await DisplayAlert("Читання", "Збережених файлів немає.", "OK");
            return;
        }

        string? choice = await DisplayActionSheet("Оберіть файл", "Скасувати", null, files);
        if (string.IsNullOrEmpty(choice) || choice == "Скасувати") return;

        try
        {
            string json = await File.ReadAllTextAsync(Path.Combine(dir, choice));
            var loaded = Spreadsheet.FromJson(json);
            loaded.RequireLogicalResult = RequireLogical;

            _sheet = loaded;
            _selected = null;
            textInput.Text = "";
            _showValues = false;
            BuildGrid();
            RefreshDisplay();
        }
        catch (Exception ex)
        {
            await DisplayAlert("Помилка", $"Не вдалося прочитати файл: {ex.Message}", "OK");
        }
    }


    async void ExitButton_Clicked(object? sender, EventArgs e)
    {
        bool answer = await DisplayAlert("Підтвердження", "Ви дійсно хочете вийти?", "Так", "Ні");
        if (answer) Application.Current?.Quit();
    }

    async void HelpButton_Clicked(object? sender, EventArgs e)
    {
        await DisplayAlert("Довідка",
            "Лабораторна робота 1. Електронна таблиця (варіант 64).\n" +
            "Виконала: Бутенко Крістіна, група к-25.\n\n" +
            "Клітинка містить вираз з: цілих чисел довільної довжини, круглих дужок, " +
            "посилань на клітинки (A1, B12) та операцій.\n\n" +
            "Операції (від найвищого пріоритету):\n" +
            "• mod, div — залишок і цілочисельне ділення;\n" +
            "• mmax(x1,...,xN), mmin(x1,...,xN), N >= 1;\n" +
            "• =, <, > — порівняння;\n" +
            "• not;\n" +
            "• and;\n" +
            "• or.\n\n" +
            "Результат виразу: логічний (true/false) або число.\n" +
            "Приклад: not (mmax(A1, 5) > 7) or B2 = 10 mod 3\n\n" +
            "Режими: ВИРАЗ — показано формули (редагування), ЗНАЧЕННЯ — результати " +
            "(лише перегляд). Перемикайте кнопкою «Режим» або натисніть «Порахувати».\n\n" +
            "Помилки (синтаксис, типи, ділення на нуль, порожня клітинка, цикл) " +
            "показуються під рядком формул із позицією в виразі.\n\n" +
            "Збереження: файл .json у папці застосунку, шлях показується після збереження.",
            "OK");
    }


    void AddRowButton_Clicked(object? sender, EventArgs e)
    {
        CommitEntries();
        _sheet.AddRow();
        int row = _sheet.RowCount;
        AddRowHeader(row);
        for (int c = 1; c <= _sheet.ColumnCount; c++) CreateEntry(row, c);
    }

    void AddColumnButton_Clicked(object? sender, EventArgs e)
    {
        CommitEntries();
        _sheet.AddColumn();
        int col = _sheet.ColumnCount;
        AddColumnHeader(col);
        for (int r = 1; r <= _sheet.RowCount; r++) CreateEntry(r, col);
    }

    async void DeleteRowButton_Clicked(object? sender, EventArgs e)
    {
        CommitEntries();
        int row = _sheet.RowCount;

        if (!_sheet.RemoveLastRow())
        {
            await DisplayAlert("Увага", "Не можна видалити останній рядок.", "OK");
            return;
        }

        for (int c = 1; c <= _sheet.ColumnCount; c++) RemoveEntry(row, c);
        grid.Children.Remove(_rowLabels[^1]);
        _rowLabels.RemoveAt(_rowLabels.Count - 1);
        grid.RowDefinitions.RemoveAt(grid.RowDefinitions.Count - 1);

        if (_selected.HasValue && _selected.Value.row == row) ClearSelection();
        RefreshDisplay();
    }

    async void DeleteColumnButton_Clicked(object? sender, EventArgs e)
    {
        CommitEntries();
        int col = _sheet.ColumnCount;

        if (!_sheet.RemoveLastColumn())
        {
            await DisplayAlert("Увага", "Не можна видалити останній стовпчик.", "OK");
            return;
        }

        for (int r = 1; r <= _sheet.RowCount; r++) RemoveEntry(r, col);
        grid.Children.Remove(_colLabels[^1]);
        _colLabels.RemoveAt(_colLabels.Count - 1);
        grid.ColumnDefinitions.RemoveAt(grid.ColumnDefinitions.Count - 1);

        if (_selected.HasValue && _selected.Value.col == col) ClearSelection();
        RefreshDisplay();
    }

    void ClearSelection()
    {
        _selected = null;
        textInput.Text = "";
    }
}