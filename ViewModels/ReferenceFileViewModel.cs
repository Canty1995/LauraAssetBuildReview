using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using LauraAssetBuildReview.Models;
using LauraAssetBuildReview.Services;
using System.Collections.ObjectModel;
using System.Windows;

namespace LauraAssetBuildReview.ViewModels;

public partial class ReferenceFileViewModel : ObservableObject
{
    private readonly PowerPointReader _powerPointReader = new();

    [ObservableProperty]
    private string _filePath = string.Empty;

    [ObservableProperty]
    private string _displayName = "Reference File";

    [ObservableProperty]
    private ReferenceFileConfig _config = new();

    [ObservableProperty]
    private ObservableCollection<SlideSelectionItem> _availableSlides = new();

    [ObservableProperty]
    private bool _isPowerPointFile = false;

    public ReferenceFileViewModel(string displayName, int priority)
    {
        DisplayName = displayName;
        Config.Priority = priority;
        Config.FileType = "Excel"; // Default to Excel
    }

    partial void OnFilePathChanged(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            IsPowerPointFile = false;
            AvailableSlides.Clear();
            Config.FileType = "Excel";
            return;
        }

        var extension = System.IO.Path.GetExtension(value).ToLowerInvariant();
        IsPowerPointFile = extension == ".pptx";
        Config.FileType = IsPowerPointFile ? "PowerPoint" : "Excel";

        if (IsPowerPointFile)
        {
            LoadSlideList();
        }
        else
        {
            AvailableSlides.Clear();
        }
    }

    private void LoadSlideList()
    {
        AvailableSlides.Clear();
        
        if (string.IsNullOrWhiteSpace(FilePath) || !System.IO.File.Exists(FilePath))
            return;

        try
        {
            var slideCount = _powerPointReader.GetSlideCount(FilePath);
            for (int i = 1; i <= slideCount; i++)
            {
                var isSelected = Config.SelectedSlides?.Contains(i) ?? false;
                AvailableSlides.Add(new SlideSelectionItem
                {
                    SlideNumber = i,
                    IsSelected = isSelected
                });
            }
        }
        catch
        {
            // If we can't read slides, just continue
        }
    }

    [RelayCommand]
    private void Browse()
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "Excel Files (*.xlsx)|*.xlsx|PowerPoint Files (*.pptx)|*.pptx|All Files (*.*)|*.*",
            Title = $"Select {DisplayName}"
        };

        if (dialog.ShowDialog() == true)
        {
            FilePath = dialog.FileName;
            Config.FilePath = FilePath;
        }
    }

    [RelayCommand]
    private void UpdateSelectedSlides()
    {
        if (Config.SelectedSlides == null)
        {
            Config.SelectedSlides = new List<int>();
        }
        else
        {
            Config.SelectedSlides.Clear();
        }

        foreach (var slide in AvailableSlides)
        {
            if (slide.IsSelected)
            {
                Config.SelectedSlides.Add(slide.SlideNumber);
            }
        }
    }

    [RelayCommand]
    private void Remove()
    {
        // This will be handled by the parent ViewModel
        if (Application.Current.MainWindow?.DataContext is MainViewModel mainVm)
        {
            mainVm.RemoveReferenceFile(this);
        }
    }
}

public class SlideSelectionItem : ObservableObject
{
    private bool _isSelected;

    public int SlideNumber { get; set; }
    
    public bool IsSelected
    {
        get => _isSelected;
        set => SetProperty(ref _isSelected, value);
    }
}

public class ManualMappingViewModel : ObservableObject
{
    private string _filePath = string.Empty;
    private string _dropdownOption = string.Empty;

    public string FilePath
    {
        get => _filePath;
        set => SetProperty(ref _filePath, value);
    }

    public string DropdownOption
    {
        get => _dropdownOption;
        set => SetProperty(ref _dropdownOption, value);
    }
}
