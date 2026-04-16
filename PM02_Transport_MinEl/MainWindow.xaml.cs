using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Excel = Microsoft.Office.Interop.Excel;

namespace PM02_Transport_MinEl
{
    public partial class MainWindow : Window
    {
        private List<List<TextBox>> matrixCells = new List<List<TextBox>>();
        private double[] supplies;
        private double[] demands;
        private double[,] savedCosts;

        private double[,] solution;
        private double totalCost;

        public MainWindow()
        {
            InitializeComponent();
            CreateTable_Click(null, null);
        }

        #region Работа с таблицей
        private void CreateTable_Click(object sender, RoutedEventArgs e)
        {
            if (!int.TryParse(SuppliersBox.Text, out int suppliers) || !int.TryParse(ConsumersBox.Text, out int consumers))
            { MessageBox.Show("Введите корректные числа"); return; }

            if (MatrixGrid.Items.Count > 0)
                SaveAllCurrentData(int.Parse(SuppliersBox.Text), int.Parse(ConsumersBox.Text));

            CreateMatrixTableWithRestoration(suppliers, consumers, Math.Min(supplies?.Length ?? suppliers, suppliers), false, 0);
        }

        private void CreateMatrixTableWithRestoration(int suppliers, int consumers, int oldDimension, bool isConsumerAdded, double distributionHint)
        {
            MatrixGrid.Items.Clear();
            ColumnHeaders.Items.Clear();
            matrixCells.Clear();

            for (int j = 0; j < consumers; j++)
            {
                var header = new TextBox
                {
                    Width = 60,
                    Height = 25,
                    Margin = new Thickness(2),
                    ToolTip = $"Потребность {j + 1}",
                    VerticalContentAlignment = VerticalAlignment.Center,
                    HorizontalContentAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Top,
                    Background = Brushes.LightBlue
                };
                header.Text = (demands != null && j < demands.Length) ? demands[j].ToString() : "0";
                AddValidation(header);
                ColumnHeaders.Items.Add(header);
            }

            for (int i = 0; i < suppliers; i++)
            {
                var rowPanel = new StackPanel { Orientation = Orientation.Horizontal };
                var resourceBox = new TextBox
                {
                    Width = 60,
                    Height = 25,
                    Margin = new Thickness(2),
                    ToolTip = $"Запас поставщика {i + 1}",
                    VerticalContentAlignment = VerticalAlignment.Center,
                    HorizontalContentAlignment = HorizontalAlignment.Center,
                    Background = Brushes.LightGray
                };
                resourceBox.Text = (supplies != null && i < supplies.Length) ? supplies[i].ToString() : "0";
                AddValidation(resourceBox);
                rowPanel.Children.Add(resourceBox);

                var rowCells = new List<TextBox>();
                for (int j = 0; j < consumers; j++)
                {
                    var cell = new TextBox
                    {
                        Width = 60,
                        Height = 25,
                        Margin = new Thickness(2),
                        ToolTip = $"Стоимость из {i + 1} в {j + 1}",
                        VerticalContentAlignment = VerticalAlignment.Center,
                        HorizontalContentAlignment = HorizontalAlignment.Center
                    };

                    if (isConsumerAdded && j >= oldDimension) { cell.Text = distributionHint.ToString("F1"); cell.Background = Brushes.LightGray; }
                    else if (!isConsumerAdded && i >= oldDimension) { cell.Text = distributionHint.ToString("F1"); }
                    else { cell.Text = (savedCosts != null && i < savedCosts.GetLength(0) && j < savedCosts.GetLength(1)) ? savedCosts[i, j].ToString() : "0"; }

                    AddValidation(cell);
                    rowPanel.Children.Add(cell);
                    rowCells.Add(cell);
                }
                MatrixGrid.Items.Add(rowPanel);
                matrixCells.Add(rowCells);
            }
        }

        private void SaveAllCurrentData(int suppliers, int consumers)
        {
            try
            {
                if (MatrixGrid.Items.Count == 0 || ColumnHeaders.Items.Count == 0) return;
                supplies = new double[suppliers]; demands = new double[consumers];
                for (int i = 0; i < suppliers && i < MatrixGrid.Items.Count; i++)
                {
                    var rowPanel = MatrixGrid.Items[i] as StackPanel;
                    if (rowPanel?.Children.Count > 0 && double.TryParse((rowPanel.Children[0] as TextBox)?.Text, out double supply))
                        supplies[i] = supply;
                }
                for (int j = 0; j < consumers && j < ColumnHeaders.Items.Count; j++)
                    if (double.TryParse((ColumnHeaders.Items[j] as TextBox)?.Text, out double demand)) demands[j] = demand;

                savedCosts = new double[suppliers, consumers];
                for (int i = 0; i < suppliers && i < matrixCells.Count; i++)
                    for (int j = 0; j < consumers && j < matrixCells[i].Count; j++)
                        double.TryParse(matrixCells[i][j].Text, out savedCosts[i, j]);
            }
            catch (Exception ex) { Console.WriteLine($"Warning: {ex.Message}"); }
        }

        private void ClearTable_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in MatrixGrid.Items)
                if (item is StackPanel rowPanel) foreach (var child in rowPanel.Children) if (child is TextBox textBox) { textBox.Text = "0"; textBox.Background = Brushes.White; }
            foreach (var item in ColumnHeaders.Items)
                if (item is TextBox textBox) { textBox.Text = "0"; textBox.Background = Brushes.White; }
        }

        private void AddValidation(TextBox textBox)
        {
            textBox.PreviewTextInput += (s, e) => { e.Handled = !IsTextAllowed(e.Text); };
            textBox.PreviewKeyDown += (s, e) => { if (e.Key == Key.Space) e.Handled = true; };
        }

        private static bool IsTextAllowed(string text)
        {
            if (string.IsNullOrEmpty(text)) return true;
            foreach (char c in text) { if (!char.IsDigit(c) && c != '.' && c != ',' && c != '-') return false; }
            return true;
        }
        #endregion

        #region Балансировка
        private void BalanceProblem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!int.TryParse(SuppliersBox.Text, out int currentSuppliers) || !int.TryParse(ConsumersBox.Text, out int currentConsumers))
                { MessageBox.Show("Некорректные размеры"); return; }
                double distributionHint = 0; double.TryParse(DistributionHintBox.Text, out distributionHint);

                double totalSupply = 0, totalDemand = 0;
                double[] currentSupplies = new double[currentSuppliers]; double[] currentDemands = new double[currentConsumers];

                for (int i = 0; i < currentSuppliers && i < MatrixGrid.Items.Count; i++)
                {
                    var rowPanel = MatrixGrid.Items[i] as StackPanel;
                    if (double.TryParse((rowPanel?.Children[0] as TextBox)?.Text, out double supply)) { currentSupplies[i] = supply; totalSupply += supply; }
                }
                for (int j = 0; j < currentConsumers && j < ColumnHeaders.Items.Count; j++)
                    if (double.TryParse((ColumnHeaders.Items[j] as TextBox)?.Text, out double demand)) { currentDemands[j] = demand; totalDemand += demand; }

                SaveAllCurrentData(currentSuppliers, currentConsumers);
                double difference = Math.Abs(totalSupply - totalDemand);

                if (difference < 0.001) { MessageBox.Show("Задача уже сбалансирована!"); return; }

                if (totalSupply > totalDemand) AddFictitiousConsumer(difference, currentSupplies, currentDemands, distributionHint);
                else AddFictitiousSupplier(difference, currentSupplies, currentDemands, distributionHint);
            }
            catch (Exception ex) { MessageBox.Show($"Ошибка: {ex.Message}"); }
        }

        private void AddFictitiousConsumer(double missingDemand, double[] currentSupplies, double[] currentDemands, double distributionHint)
        {
            int oldConsumers = currentDemands.Length; int newConsumers = oldConsumers + 1; ConsumersBox.Text = newConsumers.ToString();
            supplies = currentSupplies; demands = new double[newConsumers]; Array.Copy(currentDemands, demands, oldConsumers); demands[newConsumers - 1] = missingDemand;
            CreateMatrixTableWithRestoration(int.Parse(SuppliersBox.Text), newConsumers, oldConsumers, true, distributionHint);
            ((ColumnHeaders.Items[newConsumers - 1] as TextBox).Text) = missingDemand.ToString("F1");
        }

        private void AddFictitiousSupplier(double missingSupply, double[] currentSupplies, double[] currentDemands, double distributionHint)
        {
            int oldSuppliers = currentSupplies.Length; int newSuppliers = oldSuppliers + 1; SuppliersBox.Text = newSuppliers.ToString();
            supplies = new double[newSuppliers]; Array.Copy(currentSupplies, supplies, oldSuppliers); supplies[newSuppliers - 1] = missingSupply; demands = currentDemands;
            CreateMatrixTableWithRestoration(newSuppliers, int.Parse(ConsumersBox.Text), oldSuppliers, false, distributionHint);
            (((MatrixGrid.Items[newSuppliers - 1] as StackPanel).Children[0] as TextBox).Text) = missingSupply.ToString("F1");
        }

        private bool IsProblemBalanced()
        {
            if (!int.TryParse(SuppliersBox.Text, out int suppliers) || !int.TryParse(ConsumersBox.Text, out int consumers) || MatrixGrid.Items.Count == 0) return false;
            double totalSupply = 0, totalDemand = 0;
            for (int i = 0; i < suppliers && i < MatrixGrid.Items.Count; i++)
                if (double.TryParse(((MatrixGrid.Items[i] as StackPanel)?.Children[0] as TextBox)?.Text, out double supply)) totalSupply += supply;
            for (int j = 0; j < consumers && j < ColumnHeaders.Items.Count; j++)
                if (double.TryParse((ColumnHeaders.Items[j] as TextBox)?.Text, out double demand)) totalDemand += demand;
            return Math.Abs(totalSupply - totalDemand) < 0.001;
        }
        #endregion

        #region Решение (Минимальных элементов)
        private void Solve_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (MatrixGrid.Items.Count == 0) { MessageBox.Show("Сначала создайте таблицу!"); return; }
                if (!IsProblemBalanced()) { if (MessageBox.Show("Задача не сбалансирована! Балансировать?", "Внимание", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes) { BalanceProblem_Click(sender, e); return; } }

                int suppliersCount = matrixCells.Count; int consumersCount = matrixCells.Count > 0 ? matrixCells[0].Count : 0;
                if (suppliersCount == 0 || consumersCount == 0) { MessageBox.Show("Таблица пуста!"); return; }

                double[] sup = new double[suppliersCount]; double[] dem = new double[consumersCount]; double[,] costs = new double[suppliersCount, consumersCount];
                for (int i = 0; i < suppliersCount; i++) if (double.TryParse(((MatrixGrid.Items[i] as StackPanel)?.Children[0] as TextBox)?.Text, out double s)) sup[i] = s;
                for (int j = 0; j < consumersCount; j++) if (double.TryParse((ColumnHeaders.Items[j] as TextBox)?.Text, out double d)) dem[j] = d;
                for (int i = 0; i < suppliersCount; i++) for (int j = 0; j < consumersCount; j++) double.TryParse(matrixCells[i][j].Text, out costs[i, j]);

                SolveMinimumElements(sup, dem, costs);
                ShowSolutionResults(suppliersCount, consumersCount, "Минимальных элементов");
            }
            catch (Exception ex) { MessageBox.Show($"Ошибка: {ex.Message}"); }
        }

        private void SolveMinimumElements(double[] sup, double[] dem, double[,] costs)
        {
            int sCount = sup.Length; int cCount = dem.Length;
            solution = new double[sCount, cCount]; totalCost = 0;
            double[] remSup = (double[])sup.Clone(); double[] remDem = (double[])dem.Clone();

            while (true)
            {
                int minI = -1, minJ = -1; double minCost = double.MaxValue;
                for (int i = 0; i < sCount; i++)
                {
                    if (remSup[i] <= 0) continue;
                    for (int j = 0; j < cCount; j++)
                    {
                        if (remDem[j] <= 0) continue;
                        if (costs[i, j] < minCost) { minCost = costs[i, j]; minI = i; minJ = j; }
                    }
                }
                if (minI == -1 || minJ == -1) break;

                double amount = Math.Min(remSup[minI], remDem[minJ]);
                solution[minI, minJ] = amount;
                totalCost += amount * costs[minI, minJ];
                remSup[minI] -= amount; remDem[minJ] -= amount;
            }
        }
        #endregion

        #region Вывод результатов (Упрощенный)
        private void ShowSolutionResults(int suppliers, int consumers, string method)
        {
            Window resultsWindow = new Window
            {
                Title = $"Результаты ({method})",
                Width = 500,
                Height = 400,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                ResizeMode = ResizeMode.NoResize
            };

            StackPanel mainPanel = new StackPanel { Margin = new Thickness(15) };

            TextBlock title = new TextBlock { Text = $"Опорный план: {method}", FontSize = 16, FontWeight = FontWeights.Bold, HorizontalAlignment = HorizontalAlignment.Center, Margin = new Thickness(0, 0, 0, 15) };
            mainPanel.Children.Add(title);

            Grid resultsGrid = new Grid { Margin = new Thickness(0, 0, 0, 15) };
            for (int j = 0; j <= consumers; j++) resultsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(60) });
            for (int i = 0; i <= suppliers; i++) resultsGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(30) });

            for (int j = 0; j < consumers; j++)
            {
                TextBlock header = new TextBlock { Text = $"П{j + 1}", HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center, FontWeight = FontWeights.Bold, Background = Brushes.LightGray };
                Grid.SetRow(header, 0); Grid.SetColumn(header, j + 1); resultsGrid.Children.Add(header);
            }
            for (int i = 0; i < suppliers; i++)
            {
                TextBlock header = new TextBlock { Text = $"С{i + 1}", HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center, FontWeight = FontWeights.Bold, Background = Brushes.LightGray };
                Grid.SetRow(header, i + 1); Grid.SetColumn(header, 0); resultsGrid.Children.Add(header);
            }
            for (int i = 0; i < suppliers; i++)
            {
                for (int j = 0; j < consumers; j++)
                {
                    Border border = new Border { BorderBrush = Brushes.Black, BorderThickness = new Thickness(1), Background = solution[i, j] > 0 ? Brushes.LightGreen : Brushes.White };
                    border.Child = new TextBlock { Text = solution[i, j].ToString("F1"), HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center };
                    Grid.SetRow(border, i + 1); Grid.SetColumn(border, j + 1); resultsGrid.Children.Add(border);
                }
            }
            mainPanel.Children.Add(resultsGrid);

            TextBlock costText = new TextBlock
            {
                Text = $"Общая стоимость перевозок: {totalCost:F2}",
                FontSize = 16,
                FontWeight = FontWeights.Bold,
                HorizontalAlignment = HorizontalAlignment.Center,
                Foreground = Brushes.DarkBlue,
                Margin = new Thickness(0, 10, 0, 20)
            };
            mainPanel.Children.Add(costText);

            Button closeBtn = new Button { Content = "Закрыть", Width = 100, HorizontalAlignment = HorizontalAlignment.Center };
            closeBtn.Click += (s, e) => resultsWindow.Close();
            mainPanel.Children.Add(closeBtn);

            resultsWindow.Content = mainPanel;
            resultsWindow.ShowDialog();
        }
        #endregion

        #region Excel Import / Export
        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            if (matrixCells.Count == 0) { MessageBox.Show("Сначала создайте таблицу"); return; }
            Microsoft.Win32.SaveFileDialog sfd = new Microsoft.Win32.SaveFileDialog(); sfd.Filter = "Excel files (*.xlsx)|*.xlsx";
            if (sfd.ShowDialog() != true) return;

            Excel.Application xlApp = null; Excel.Workbook xlWorkbook = null; Excel.Worksheet xlWorksheet = null;
            try
            {
                xlApp = new Excel.Application(); xlApp.DisplayAlerts = false;
                xlWorkbook = xlApp.Workbooks.Add(); xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[1];

                int suppliers = matrixCells.Count; int consumers = matrixCells[0].Count;

                for (int j = 0; j < consumers; j++) xlWorksheet.Cells[1, j + 2] = (ColumnHeaders.Items[j] as TextBox)?.Text;

                for (int i = 0; i < suppliers; i++)
                {
                    xlWorksheet.Cells[i + 2, 1] = ((MatrixGrid.Items[i] as StackPanel)?.Children[0] as TextBox)?.Text;
                    for (int j = 0; j < consumers; j++) xlWorksheet.Cells[i + 2, j + 2] = matrixCells[i][j].Text;
                }
                xlWorksheet.Columns.AutoFit();
                xlWorkbook.SaveAs(sfd.FileName);
                MessageBox.Show("Успешно экспортировано!");
            }
            catch (Exception ex) { MessageBox.Show($"Ошибка экспорта: {ex.Message}"); }
            finally { ReleaseExcelObject(null, xlWorksheet, xlWorkbook, xlApp); xlWorksheet = null; xlWorkbook = null; xlApp = null; }
        }

        private void ImportFromExcel_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog ofd = new Microsoft.Win32.OpenFileDialog(); ofd.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls";
            if (ofd.ShowDialog() != true) return;

            Excel.Application xlApp = null; Excel.Workbook xlWorkbook = null; Excel.Worksheet xlWorksheet = null; Excel.Range xlRange = null;
            try
            {
                xlApp = new Excel.Application(); xlWorkbook = xlApp.Workbooks.Open(ofd.FileName); xlWorksheet = xlWorkbook.Sheets[1]; xlRange = xlWorksheet.UsedRange;
                int rowCount = xlRange.Rows.Count; int colCount = xlRange.Columns.Count;

                if (rowCount < 2 || colCount < 2) { MessageBox.Show("Файл слишком пустой."); return; }

                int excelSuppliers = rowCount - 1;
                int excelConsumers = colCount - 1;

                SuppliersBox.Text = excelSuppliers.ToString();
                ConsumersBox.Text = excelConsumers.ToString();
                CreateTable_Click(null, null);

                for (int j = 0; j < excelConsumers; j++)
                {
                    string val = xlRange.Cells[1, j + 2].Value?.ToString();
                    if (string.IsNullOrWhiteSpace(val)) (ColumnHeaders.Items[j] as TextBox).Text = "0";
                    else
                    {
                        if (double.TryParse(val.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double num))
                            (ColumnHeaders.Items[j] as TextBox).Text = num.ToString();
                        else throw new Exception($"Нечисловое значение потребителя в столбце {j + 2}");
                    }
                }

                for (int i = 0; i < excelSuppliers; i++)
                {
                    string supVal = xlRange.Cells[i + 2, 1].Value?.ToString();
                    if (string.IsNullOrWhiteSpace(supVal)) ((MatrixGrid.Items[i] as StackPanel).Children[0] as TextBox).Text = "0";
                    else
                    {
                        if (double.TryParse(supVal.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double supNum))
                            ((MatrixGrid.Items[i] as StackPanel).Children[0] as TextBox).Text = supNum.ToString();
                        else throw new Exception($"Нечисловое значение поставщика в строке {i + 2}");
                    }

                    for (int j = 0; j < excelConsumers; j++)
                    {
                        string costVal = xlRange.Cells[i + 2, j + 2].Value?.ToString();
                        if (string.IsNullOrWhiteSpace(costVal)) matrixCells[i][j].Text = "0";
                        else
                        {
                            if (double.TryParse(costVal.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double costNum))
                                matrixCells[i][j].Text = costNum.ToString();
                            else throw new Exception($"Нечисловое значение тарифа: строка {i + 2}, столбец {j + 2}");
                        }
                    }
                }
                MessageBox.Show("Данные успешно импортированы!");
            }
            catch (Exception ex) { MessageBox.Show($"Ошибка импорта:\n{ex.Message}"); }
            finally { ReleaseExcelObject(xlRange, xlWorksheet, xlWorkbook, xlApp); xlRange = null; xlWorksheet = null; xlWorkbook = null; xlApp = null; }
        }

        private void ReleaseExcelObject(object xlRange = null, object xlWorksheet = null, object xlWorkbook = null, object xlApp = null)
        {
            try
            {
                if (xlRange != null) { Marshal.ReleaseComObject(xlRange); xlRange = null; }
                if (xlWorksheet != null) { Marshal.ReleaseComObject(xlWorksheet); xlWorksheet = null; }
                if (xlWorkbook != null) { ((Excel.Workbook)xlWorkbook).Close(false); Marshal.ReleaseComObject(xlWorkbook); xlWorkbook = null; }
                if (xlApp != null) { ((Excel.Application)xlApp).Quit(); Marshal.ReleaseComObject(xlApp); xlApp = null; }
            }
            catch { }
            finally { GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); GC.WaitForPendingFinalizers(); }
        }
        #endregion
    }
}