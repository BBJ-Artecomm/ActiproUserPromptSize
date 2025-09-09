using ActiproSoftware.Windows.Controls;
using ActiproSoftware.Windows.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ActiproUserPromptSize
{
    /// <summary>
    /// Logique d'interaction pour MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            //
            // SAMPLE: Customize the header and content
            //

            var statusText = new TextBlock()
            {
                Text = "Estimated time remaining:",
                Margin = new Thickness(0, 2, 0, 2)
            };

            var progressBar = new AnimatedProgressBar()
            {
                Margin = new Thickness(0, 5, 0, 0),
                Minimum = 0,
                Maximum = 100,
                Value = 0,
                Height = 20,
                State = OperationState.Normal
            };

            UserPromptBuilder.Configure()
                .WithHeaderContent("Exporting Project (Sample Project)")
                .WithHeaderForeground(Colors.White)
                .WithHeaderBackground(
                    new LinearGradientBrush(
                        startColor: (Color)ColorConverter.ConvertFromString("#094C75"),
                        endColor: (Color)ColorConverter.ConvertFromString("#066F5C"),
                        angle: 0d)
                )
                //.WithStatusImage(ImageLoader.GetIcon("Save32.png"))
                .WithStandardButtons(UserPromptStandardButtons.Cancel)
                .WithContent(new StackPanel()
                {
                    Children = {
            new TextBlock() {
                Inlines = {
                    new Run("to "),
                    new Run("Project Templates") { FontWeight = FontWeights.Bold },
                    new Run(@" (C:\Templates\ProjectTemplates)"),
                }
            },
            statusText,
            progressBar
                    }
                })
                .WithCheckBoxContent("Check this box to simulate an exception")
                .WithWindowStartupLocation(WindowStartupLocation.CenterOwner)
                .WithAutoSize(true, minimumWidth: 1400)
                .BeforeShow(builder => {
                    // Do work here
                    Task.Run(() => {
                        var totalTime = TimeSpan.FromSeconds(10);
                        var startTaskTime = DateTime.Now;
                        var isCompleted = false;
                        var throwException = false;
                        try
                        {
                            do
                            {
                                if (throwException)
                                    throw new ApplicationException("An error was encountered during export.");
                                Thread.Sleep(100);
                                var elapsedTime = DateTime.Now - startTaskTime;
                                var remainingTime = totalTime - elapsedTime;
                                var percentageComplete = (((double)elapsedTime.Ticks / (double)totalTime.Ticks) * 100);
                                Dispatcher.InvokeAsync(() => {
                                    throwException = builder.Instance?.IsChecked == true;
                                    if (!throwException)
                                    {
                                        progressBar.Value = percentageComplete;
                                        statusText.Text = $"Estimated time remaining: {remainingTime.TotalSeconds.Round(RoundMode.Ceiling)} seconds";
                                        isCompleted = (progressBar.IsCompleted || (builder.Instance?.Result != null));
                                    }
                                });
                            } while (!isCompleted);
                            if (isCompleted)
                            {
                                Dispatcher.InvokeAsync(() => {
                                    if ((builder.Instance != null) && (builder.Instance.Result is null))
                                        builder.Instance.Result = builder.Instance.DefaultResult;
                                });
                            }
                        }
                        catch (Exception ex)
                        {
                            Dispatcher.InvokeAsync(() => {
                                statusText.Text = "Error: " + ex.Message;
                                progressBar.State = OperationState.Error;
                            });
                        }
                    });
                })
                .Show();

        }
    }
}
