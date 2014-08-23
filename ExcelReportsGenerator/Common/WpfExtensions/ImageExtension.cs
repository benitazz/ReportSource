using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;

namespace ExcelReportsGenerator.Common.WpfExtensions
{
  /// <summary>
  ///   The Image Extension class.
  /// </summary>
  public class ImageExtension: Image
  {
    #region Static Fields

    /// <summary>
    ///   The GIF source property
    /// </summary>
    public static readonly DependencyProperty GifSourceProperty = DependencyProperty.Register(
      "GifSource",
      typeof(string),
      typeof(ImageExtension),
      new UIPropertyMetadata(string.Empty, GifSourcePropertyChanged));

    /// <summary>
    ///   The frame index property
    /// </summary>
    public static readonly DependencyProperty FrameIndexProperty = DependencyProperty.Register(
      "FrameIndex",
      typeof(int),
      typeof(ImageExtension),
      new UIPropertyMetadata(0, ChangingFrameIndex));

    /// <summary>
    ///   The automatic start property
    /// </summary>
    public static readonly DependencyProperty AutoStartProperty = DependencyProperty.Register(
      "AutoStart",
      typeof(bool),
      typeof(ImageExtension),
      new UIPropertyMetadata(false, AutoStartPropertyChanged));
    
    #endregion

    #region Fields

    /// <summary>
    /// The _animation.
    /// </summary>
    private Int32Animation _animation;

    /// <summary>
    /// The _gif decoder.
    /// </summary>
    private GifBitmapDecoder _gifDecoder;

    /// <summary>
    /// The _is initialized.
    /// </summary>
    private bool _isInitialized;

    #endregion

    #region Constructors and Destructors

    /// <summary>
    /// Initializes static members of the <see cref="ImageExtension"/> class. 
    ///   GIFs the image.
    /// </summary>
    static ImageExtension()
    {
      VisibilityProperty.OverrideMetadata(
        typeof(ImageExtension),
        new FrameworkPropertyMetadata(VisibilityPropertyChanged));
    }

    #endregion

    #region Public Properties

    /// <summary>
    ///   Defines whether the animation starts on it's own
    /// </summary>
    public bool AutoStart
    {
      get
      {
        return (bool)this.GetValue(AutoStartProperty);
      }

      set
      {
        this.SetValue(AutoStartProperty, value);
      }
    }

    /// <summary>
    ///   Gets or sets the index of the frame.
    /// </summary>
    /// <value>
    ///   The index of the frame.
    /// </value>
    public int FrameIndex
    {
      get
      {
        return (int)this.GetValue(FrameIndexProperty);
      }

      set
      {
        this.SetValue(FrameIndexProperty, value);
      }
    }

    /// <summary>
    ///   Gets or sets the GIF source.
    /// </summary>
    /// <value>
    ///   The GIF source.
    /// </value>
    public string GifSource
    {
      get
      {
        return (string)this.GetValue(GifSourceProperty);
      }

      set
      {
        this.SetValue(GifSourceProperty, value);
      }
    }

    #endregion

    #region Public Methods and Operators

    /// <summary>
    ///   Starts the animation
    /// </summary>
    public void StartAnimation()
    {
      if (!this._isInitialized)
      {
        this.Initialize();
      }

      this.BeginAnimation(FrameIndexProperty, this._animation);
    }

    /// <summary>
    ///   Stops the animation
    /// </summary>
    public void StopAnimation()
    {
      this.BeginAnimation(FrameIndexProperty, null);
    }

    #endregion

    #region Methods

    /// <summary>
    /// Automatics the start property changed.
    /// </summary>
    /// <param name="sender">
    /// The sender.
    /// </param>
    /// <param name="e">
    /// The <see cref="System.Windows.DependencyPropertyChangedEventArgs"/> instance containing the event
    ///   data.
    /// </param>
    private static void AutoStartPropertyChanged(DependencyObject sender, DependencyPropertyChangedEventArgs e)
    {
      if (!(bool)e.NewValue)
      {
        return;
      }

      var imageExtension = sender as ImageExtension;

      if (imageExtension != null)
      {
        imageExtension.StartAnimation();
      }
    }

    /// <summary>
    /// The changing frame index.
    /// </summary>
    /// <param name="obj">
    /// The obj.
    /// </param>
    /// <param name="ev">
    /// The ev.
    /// </param>
    private static void ChangingFrameIndex(DependencyObject obj, DependencyPropertyChangedEventArgs ev)
    {
      var gifImage = obj as ImageExtension;

      if (gifImage != null)
      {
        gifImage.Source = gifImage._gifDecoder.Frames[(int)ev.NewValue];
      }
    }

    /// <summary>
    /// GIFs the source property changed.
    /// </summary>
    /// <param name="sender">
    /// The sender.
    /// </param>
    /// <param name="e">
    /// The <see cref="System.Windows.DependencyPropertyChangedEventArgs"/> instance containing the event
    ///   data.
    /// </param>
    private static void GifSourcePropertyChanged(DependencyObject sender, DependencyPropertyChangedEventArgs e)
    {
      var imageExtension = sender as ImageExtension;

      if (imageExtension != null)
      {
        imageExtension.Initialize();
      }
    }

    /// <summary>
    /// Visibilities the property changed.
    /// </summary>
    /// <param name="sender">
    /// The sender.
    /// </param>
    /// <param name="e">
    /// The <see cref="System.Windows.DependencyPropertyChangedEventArgs"/> instance containing the event
    ///   data.
    /// </param>
    private static void VisibilityPropertyChanged(DependencyObject sender, DependencyPropertyChangedEventArgs e)
    {
      if ((Visibility)e.NewValue == Visibility.Visible)
      {
        ((ImageExtension)sender).StartAnimation();
        return;
      }
     
      ((ImageExtension)sender).StopAnimation();
    }

    /// <summary>
    ///   Initializes this instance.
    /// </summary>
    private void Initialize()
    {
      /*if (string.IsNullOrEmpty(this.GifSource))
      {
        this.GifSource = "/ExcelReportsGenerator;component/Resources/Images/Windows8loader.gif";
      }*/

      this._gifDecoder = new GifBitmapDecoder(
        new Uri("pack://application:,,," + this.GifSource),
        BitmapCreateOptions.PreservePixelFormat,
        BitmapCacheOption.Default);

      this._animation = new Int32Animation(
        0,
        this._gifDecoder.Frames.Count - 1,
        new Duration(
          new TimeSpan(
            0,
            0,
            0,
            this._gifDecoder.Frames.Count / 10,
            (int)((this._gifDecoder.Frames.Count / 10.0 - this._gifDecoder.Frames.Count / 10) * 1000))))
      {
        RepeatBehavior = RepeatBehavior.Forever
      };

      this.Source = this._gifDecoder.Frames[0];

      this._isInitialized = true;
    }

    #endregion
  }
}
