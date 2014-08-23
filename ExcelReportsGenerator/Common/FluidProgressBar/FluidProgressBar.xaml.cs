#region

using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Media.Animation;

#endregion

namespace ExcelReportsGenerator.Common.FluidProgressBar
{
  /// <summary>
  ///   Interaction logic for FluidProgressBar.xaml
  /// </summary>
  public partial class FluidProgressBar : IDisposable
  {
    #region Static Fields

    /// <summary>
    ///   Delay Dependency Property
    /// </summary>
    public static readonly DependencyProperty DelayProperty = DependencyProperty.Register(
      "Delay", 
      typeof(Duration), 
      typeof(FluidProgressBar), 
      new FrameworkPropertyMetadata(
        new Duration(TimeSpan.FromMilliseconds(100)), 
        new PropertyChangedCallback(OnDelayChanged)));

    /// <summary>
    ///   DotHeight Dependency Property
    /// </summary>
    public static readonly DependencyProperty DotHeightProperty = DependencyProperty.Register(
      "DotHeight", 
      typeof(double), 
      typeof(FluidProgressBar), 
      new FrameworkPropertyMetadata(4.0, new PropertyChangedCallback(OnDotHeightChanged)));

    /// <summary>
    ///   DotRadiusX Dependency Property
    /// </summary>
    public static readonly DependencyProperty DotRadiusXProperty = DependencyProperty.Register(
      "DotRadiusX", 
      typeof(double), 
      typeof(FluidProgressBar), 
      new FrameworkPropertyMetadata(0.0, new PropertyChangedCallback(OnDotRadiusXChanged)));

    /// <summary>
    ///   DotRadiusY Dependency Property
    /// </summary>
    public static readonly DependencyProperty DotRadiusYProperty = DependencyProperty.Register(
      "DotRadiusY", 
      typeof(double), 
      typeof(FluidProgressBar), 
      new FrameworkPropertyMetadata(0.0, new PropertyChangedCallback(OnDotRadiusYChanged)));

    /// <summary>
    ///   DotWidth Dependency Property
    /// </summary>
    public static readonly DependencyProperty DotWidthProperty = DependencyProperty.Register(
      "DotWidth", 
      typeof(double), 
      typeof(FluidProgressBar), 
      new FrameworkPropertyMetadata(4.0, new PropertyChangedCallback(OnDotWidthChanged)));

    /// <summary>
    ///   DurationA Dependency Property
    /// </summary>
    public static readonly DependencyProperty DurationAProperty = DependencyProperty.Register(
      "DurationA", 
      typeof(Duration), 
      typeof(FluidProgressBar), 
      new FrameworkPropertyMetadata(
        new Duration(TimeSpan.FromSeconds(0.5)), 
        new PropertyChangedCallback(OnDurationAChanged)));

    /// <summary>
    ///   DurationB Dependency Property
    /// </summary>
    public static readonly DependencyProperty DurationBProperty = DependencyProperty.Register(
      "DurationB", 
      typeof(Duration), 
      typeof(FluidProgressBar), 
      new FrameworkPropertyMetadata(
        new Duration(TimeSpan.FromSeconds(1.5)), 
        new PropertyChangedCallback(OnDurationBChanged)));

    /// <summary>
    ///   DurationC Dependency Property
    /// </summary>
    public static readonly DependencyProperty DurationCProperty = DependencyProperty.Register(
      "DurationC", 
      typeof(Duration), 
      typeof(FluidProgressBar), 
      new FrameworkPropertyMetadata(
        new Duration(TimeSpan.FromSeconds(0.5)), 
        new PropertyChangedCallback(OnDurationCChanged)));

    /// <summary>
    ///   KeyFrameA Dependency Property
    /// </summary>
    public static readonly DependencyProperty KeyFrameAProperty = DependencyProperty.Register(
      "KeyFrameA", 
      typeof(double), 
      typeof(FluidProgressBar), 
      new FrameworkPropertyMetadata(0.33, new PropertyChangedCallback(OnKeyFrameAChanged)));

    /// <summary>
    ///   KeyFrameB Dependency Property
    /// </summary>
    public static readonly DependencyProperty KeyFrameBProperty = DependencyProperty.Register(
      "KeyFrameB", 
      typeof(double), 
      typeof(FluidProgressBar), 
      new FrameworkPropertyMetadata(0.63, new PropertyChangedCallback(OnKeyFrameBChanged)));

    /// <summary>
    ///   Oscillate Dependency Property
    /// </summary>
    public static readonly DependencyProperty OscillateProperty = DependencyProperty.Register(
      "Oscillate", 
      typeof(bool), 
      typeof(FluidProgressBar), 
      new FrameworkPropertyMetadata(false, new PropertyChangedCallback(OnOscillateChanged)));

    /// <summary>
    ///   ReverseDuration Dependency Property
    /// </summary>
    public static readonly DependencyProperty ReverseDurationProperty = DependencyProperty.Register(
      "ReverseDuration", 
      typeof(Duration), 
      typeof(FluidProgressBar), 
      new FrameworkPropertyMetadata(
        new Duration(TimeSpan.FromSeconds(2.9)), 
        new PropertyChangedCallback(OnReverseDurationChanged)));

    /// <summary>
    ///   TotalDuration Dependency Property
    /// </summary>
    public static readonly DependencyProperty TotalDurationProperty = DependencyProperty.Register(
      "TotalDuration", 
      typeof(Duration), 
      typeof(FluidProgressBar), 
      new FrameworkPropertyMetadata(
        new Duration(TimeSpan.FromSeconds(4.4)), 
        new PropertyChangedCallback(OnTotalDurationChanged)));

    #endregion

    #region Fields

    /// <summary>
    ///   The is storyboard running
    /// </summary>
    private bool isStoryboardRunning;

    /// <summary>
    ///   The key frame map
    /// </summary>
    private Dictionary<int, KeyFrameDetails> keyFrameMap = null;

    /// <summary>
    ///   The op key frame map
    /// </summary>
    private Dictionary<int, KeyFrameDetails> opKeyFrameMap = null;

    /// <summary>
    ///   The sb
    /// </summary>
    private Storyboard sb;

    #endregion

    #region Constructors and Destructors

    /// <summary>
    /// Initializes a new instance of the <see cref="FluidProgressBar"/> class. 
    ///   Ctor
    /// </summary>
    public FluidProgressBar()
    {
      this.InitializeComponent();

      this.keyFrameMap = new Dictionary<int, KeyFrameDetails>();
      this.opKeyFrameMap = new Dictionary<int, KeyFrameDetails>();

      this.GetKeyFramesFromStoryboard();

      this.SizeChanged += new SizeChangedEventHandler(this.OnSizeChanged);
      this.Loaded += new RoutedEventHandler(this.OnLoaded);
      this.IsVisibleChanged += new DependencyPropertyChangedEventHandler(this.OnIsVisibleChanged);
    }

    /// <summary>
    /// Finalizes an instance of the <see cref="FluidProgressBar"/> class. 
    ///   Releases unmanaged resources before an instance of the FluidProgressBar class is reclaimed by garbage collection.
    /// </summary>
    /// <remarks>
    /// NOTE: Leave out the finalizer altogether if this class doesn't own unmanaged resources itself,
    ///   but leave the other methods exactly as they are.
    ///   This method releases unmanaged resources by calling the virtual Dispose(bool), passing in 'false'.
    /// </remarks>
    ~FluidProgressBar()
    {
      this.Dispose(false);
    }

    #endregion

    #region Public Properties

    /// <summary>
    ///   Gets or sets the Delay property. This dependency property
    ///   indicates the delay between adjacent animation timelines.
    /// </summary>
    public Duration Delay
    {
      get
      {
        return (Duration)this.GetValue(DelayProperty);
      }

      set
      {
        this.SetValue(DelayProperty, value);
      }
    }

    /// <summary>
    ///   Gets or sets the DotHeight property. This dependency property
    ///   indicates the height of each of the dots.
    /// </summary>
    public double DotHeight
    {
      get
      {
        return (double)this.GetValue(DotHeightProperty);
      }

      set
      {
        this.SetValue(DotHeightProperty, value);
      }
    }

    /// <summary>
    ///   Gets or sets the DotRadiusX property. This dependency property
    ///   indicates the corner radius width of each of the dot.
    /// </summary>
    public double DotRadiusX
    {
      get
      {
        return (double)this.GetValue(DotRadiusXProperty);
      }

      set
      {
        this.SetValue(DotRadiusXProperty, value);
      }
    }

    /// <summary>
    ///   Gets or sets the DotRadiusY property. This dependency property
    ///   indicates the corner height of each of the dots.
    /// </summary>
    public double DotRadiusY
    {
      get
      {
        return (double)this.GetValue(DotRadiusYProperty);
      }

      set
      {
        this.SetValue(DotRadiusYProperty, value);
      }
    }

    /// <summary>
    ///   Gets or sets the DotWidth property. This dependency property
    ///   indicates the width of each of the dots.
    /// </summary>
    public double DotWidth
    {
      get
      {
        return (double)this.GetValue(DotWidthProperty);
      }

      set
      {
        this.SetValue(DotWidthProperty, value);
      }
    }

    /// <summary>
    ///   Gets or sets the DurationA property. This dependency property
    ///   indicates the duration of the animation from the start point till KeyFrameA.
    /// </summary>
    public Duration DurationA
    {
      get
      {
        return (Duration)this.GetValue(DurationAProperty);
      }

      set
      {
        this.SetValue(DurationAProperty, value);
      }
    }

    /// <summary>
    ///   Gets or sets the DurationB property. This dependency property
    ///   indicates the duration of the animation from the KeyFrameA till KeyFrameB.
    /// </summary>
    public Duration DurationB
    {
      get
      {
        return (Duration)this.GetValue(DurationBProperty);
      }

      set
      {
        this.SetValue(DurationBProperty, value);
      }
    }

    /// <summary>
    ///   Gets or sets the DurationC property. This dependency property
    ///   indicates the duration of the animation from KeyFrameB till the end point.
    /// </summary>
    public Duration DurationC
    {
      get
      {
        return (Duration)this.GetValue(DurationCProperty);
      }

      set
      {
        this.SetValue(DurationCProperty, value);
      }
    }

    /// <summary>
    ///   Gets or sets the KeyFrameA property. This dependency property
    ///   indicates the first KeyFrame position after the initial keyframe.
    /// </summary>
    public double KeyFrameA
    {
      get
      {
        return (double)this.GetValue(KeyFrameAProperty);
      }

      set
      {
        this.SetValue(KeyFrameAProperty, value);
      }
    }

    /// <summary>
    ///   Gets or sets the KeyFrameB property. This dependency property
    ///   indicates the second KeyFrame position after the initial keyframe.
    /// </summary>
    public double KeyFrameB
    {
      get
      {
        return (double)this.GetValue(KeyFrameBProperty);
      }

      set
      {
        this.SetValue(KeyFrameBProperty, value);
      }
    }

    /// <summary>
    ///   Gets or sets the Oscillate property. This dependency property
    ///   indicates whether the animation should oscillate.
    /// </summary>
    public bool Oscillate
    {
      get
      {
        return (bool)this.GetValue(OscillateProperty);
      }

      set
      {
        this.SetValue(OscillateProperty, value);
      }
    }

    /// <summary>
    ///   Gets or sets the ReverseDuration property. This dependency property
    ///   indicates the duration of the total animation in reverse.
    /// </summary>
    public Duration ReverseDuration
    {
      get
      {
        return (Duration)this.GetValue(ReverseDurationProperty);
      }

      set
      {
        this.SetValue(ReverseDurationProperty, value);
      }
    }

    /// <summary>
    ///   Gets or sets the TotalDuration property. This dependency property
    ///   indicates the duration of the complete animation.
    /// </summary>
    public Duration TotalDuration
    {
      get
      {
        return (Duration)this.GetValue(TotalDurationProperty);
      }

      set
      {
        this.SetValue(TotalDurationProperty, value);
      }
    }

    #endregion

    #region Public Methods and Operators

    /// <summary>
    ///   Releases all resources used by an instance of the FluidProgressBar class.
    /// </summary>
    /// <remarks>
    ///   This method calls the virtual Dispose(bool) method, passing in 'true', and then suppresses
    ///   finalization of the instance.
    /// </remarks>
    public void Dispose()
    {
      this.Dispose(true);
      GC.SuppressFinalize(this);
    }

    #endregion

    #region Methods

    /// <summary>
    /// Releases the unmanaged resources used by an instance of the FluidProgressBar class and optionally releases the
    ///   managed resources.
    /// </summary>
    /// <param name="disposing">
    /// 'true' to release both managed and unmanaged resources; 'false' to release only unmanaged
    ///   resources.
    /// </param>
    protected virtual void Dispose(bool disposing)
    {
      if (disposing)
      {
        // free managed resources here
        this.SizeChanged -= this.OnSizeChanged;
        this.Loaded -= this.OnLoaded;
        this.IsVisibleChanged -= this.OnIsVisibleChanged;
      }

      // free native resources if there are any.			
    }

    /// <summary>
    /// Provides derived classes an opportunity to handle changes to the Delay property.
    /// </summary>
    /// <param name="oldDelay">
    /// Old Value
    /// </param>
    /// <param name="newDelay">
    /// New Value
    /// </param>
    protected virtual void OnDelayChanged(Duration oldDelay, Duration newDelay)
    {
      bool isActive = this.isStoryboardRunning;
      if (isActive)
      {
        this.StopFluidAnimation();
      }

      this.UpdateTimelineDelay(newDelay);

      if (isActive)
      {
        this.StartFluidAnimation();
      }
    }

    /// <summary>
    /// Provides derived classes an opportunity to handle changes to the DotHeight property.
    /// </summary>
    /// <param name="oldDotHeight">
    /// Old Value
    /// </param>
    /// <param name="newDotHeight">
    /// New Value
    /// </param>
    protected virtual void OnDotHeightChanged(double oldDotHeight, double newDotHeight)
    {
      if (this.isStoryboardRunning)
      {
        this.RestartStoryboardAnimation();
      }
    }

    /// <summary>
    /// Provides derived classes an opportunity to handle changes to the DotRadiusX property.
    /// </summary>
    /// <param name="oldDotRadiusX">
    /// Old Value
    /// </param>
    /// <param name="newDotRadiusX">
    /// New Value
    /// </param>
    protected virtual void OnDotRadiusXChanged(double oldDotRadiusX, double newDotRadiusX)
    {
      if (this.isStoryboardRunning)
      {
        this.RestartStoryboardAnimation();
      }
    }

    /// <summary>
    /// Provides derived classes an opportunity to handle changes to the DotRadiusY property.
    /// </summary>
    /// <param name="oldDotRadiusY">
    /// Old Value
    /// </param>
    /// <param name="newDotRadiusY">
    /// New Value
    /// </param>
    protected virtual void OnDotRadiusYChanged(double oldDotRadiusY, double newDotRadiusY)
    {
      if (this.isStoryboardRunning)
      {
        this.RestartStoryboardAnimation();
      }
    }

    /// <summary>
    /// Provides derived classes an opportunity to handle changes to the DotWidth property.
    /// </summary>
    /// <param name="oldDotWidth">
    /// Old Value
    /// </param>
    /// <param name="newDotWidth">
    /// New Value
    /// </param>
    protected virtual void OnDotWidthChanged(double oldDotWidth, double newDotWidth)
    {
      if (this.isStoryboardRunning)
      {
        this.RestartStoryboardAnimation();
      }
    }

    /// <summary>
    /// Provides derived classes an opportunity to handle changes to the DurationA property.
    /// </summary>
    /// <param name="oldDurationA">
    /// Old Value
    /// </param>
    /// <param name="newDurationA">
    /// New Value
    /// </param>
    protected virtual void OnDurationAChanged(Duration oldDurationA, Duration newDurationA)
    {
      bool isActive = this.isStoryboardRunning;
      if (isActive)
      {
        this.StopFluidAnimation();
      }

      this.UpdateKeyTimes(1, newDurationA);

      if (isActive)
      {
        this.StartFluidAnimation();
      }
    }

    /// <summary>
    /// Provides derived classes an opportunity to handle changes to the DurationB property.
    /// </summary>
    /// <param name="oldDurationB">
    /// Old Value
    /// </param>
    /// <param name="newDurationB">
    /// New Value
    /// </param>
    protected virtual void OnDurationBChanged(Duration oldDurationB, Duration newDurationB)
    {
      bool isActive = this.isStoryboardRunning;
      if (isActive)
      {
        this.StopFluidAnimation();
      }

      this.UpdateKeyTimes(2, newDurationB);

      if (isActive)
      {
        this.StartFluidAnimation();
      }
    }

    /// <summary>
    /// Provides derived classes an opportunity to handle changes to the DurationC property.
    /// </summary>
    /// <param name="oldDurationC">
    /// Old Value
    /// </param>
    /// <param name="newDurationC">
    /// New Value
    /// </param>
    protected virtual void OnDurationCChanged(Duration oldDurationC, Duration newDurationC)
    {
      bool isActive = this.isStoryboardRunning;
      if (isActive)
      {
        this.StopFluidAnimation();
      }

      this.UpdateKeyTimes(3, newDurationC);

      if (isActive)
      {
        this.StartFluidAnimation();
      }
    }

    /// <summary>
    /// Provides derived classes an opportunity to handle changes to the KeyFrameA property.
    /// </summary>
    /// <param name="oldKeyFrameA">
    /// Old Value
    /// </param>
    /// <param name="newKeyFrameA">
    /// New Value
    /// </param>
    protected virtual void OnKeyFrameAChanged(double oldKeyFrameA, double newKeyFrameA)
    {
      this.RestartStoryboardAnimation();
    }

    /// <summary>
    /// Provides derived classes an opportunity to handle changes to the KeyFrameB property.
    /// </summary>
    /// <param name="oldKeyFrameB">
    /// Old Value
    /// </param>
    /// <param name="newKeyFrameB">
    /// New Value
    /// </param>
    protected virtual void OnKeyFrameBChanged(double oldKeyFrameB, double newKeyFrameB)
    {
      this.RestartStoryboardAnimation();
    }

    /// <summary>
    /// Provides derived classes an opportunity to handle changes to the Oscillate property.
    /// </summary>
    /// <param name="oldOscillate">
    /// Old Value
    /// </param>
    /// <param name="newOscillate">
    /// New Value
    /// </param>
    protected virtual void OnOscillateChanged(bool oldOscillate, bool newOscillate)
    {
      if (this.sb != null)
      {
        this.StopFluidAnimation();
        this.sb.AutoReverse = newOscillate;
        this.sb.Duration = newOscillate ? this.ReverseDuration : this.TotalDuration;
        this.StartFluidAnimation();
      }
    }

    /// <summary>
    /// Provides derived classes an opportunity to handle changes to the ReverseDuration property.
    /// </summary>
    /// <param name="oldReverseDuration">
    /// Old Value
    /// </param>
    /// <param name="newReverseDuration">
    /// New Value
    /// </param>
    protected virtual void OnReverseDurationChanged(Duration oldReverseDuration, Duration newReverseDuration)
    {
      if ((this.sb != null) && this.Oscillate)
      {
        this.sb.Duration = newReverseDuration;
        this.RestartStoryboardAnimation();
      }
    }

    /// <summary>
    /// Provides derived classes an opportunity to handle changes to the TotalDuration property.
    /// </summary>
    /// <param name="oldTotalDuration">
    /// Old Value
    /// </param>
    /// <param name="newTotalDuration">
    /// New Value
    /// </param>
    protected virtual void OnTotalDurationChanged(Duration oldTotalDuration, Duration newTotalDuration)
    {
      if ((this.sb != null) && (!this.Oscillate))
      {
        this.sb.Duration = newTotalDuration;
        this.RestartStoryboardAnimation();
      }
    }

    /// <summary>
    /// Handles changes to the Delay property.
    /// </summary>
    /// <param name="d">
    /// FluidProgressBar
    /// </param>
    /// <param name="e">
    /// DependencyProperty changed event arguments
    /// </param>
    private static void OnDelayChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
      FluidProgressBar pBar = (FluidProgressBar)d;
      Duration oldDelay = (Duration)e.OldValue;
      Duration newDelay = pBar.Delay;
      pBar.OnDelayChanged(oldDelay, newDelay);
    }

    /// <summary>
    /// Handles changes to the DotHeight property.
    /// </summary>
    /// <param name="d">
    /// FluidProgressBar
    /// </param>
    /// <param name="e">
    /// DependencyProperty changed event arguments
    /// </param>
    private static void OnDotHeightChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
      FluidProgressBar pBar = (FluidProgressBar)d;
      double oldDotHeight = (double)e.OldValue;
      double newDotHeight = pBar.DotHeight;
      pBar.OnDotHeightChanged(oldDotHeight, newDotHeight);
    }

    /// <summary>
    /// Handles changes to the DotRadiusX property.
    /// </summary>
    /// <param name="d">
    /// FluidProgressBar
    /// </param>
    /// <param name="e">
    /// DependencyProperty changed event arguments
    /// </param>
    private static void OnDotRadiusXChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
      FluidProgressBar pBar = (FluidProgressBar)d;
      double oldDotRadiusX = (double)e.OldValue;
      double newDotRadiusX = pBar.DotRadiusX;
      pBar.OnDotRadiusXChanged(oldDotRadiusX, newDotRadiusX);
    }

    /// <summary>
    /// Handles changes to the DotRadiusY property.
    /// </summary>
    /// <param name="d">
    /// FluidProgressBar
    /// </param>
    /// <param name="e">
    /// DependencyProperty changed event arguments
    /// </param>
    private static void OnDotRadiusYChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
      FluidProgressBar pBar = (FluidProgressBar)d;
      double oldDotRadiusY = (double)e.OldValue;
      double newDotRadiusY = pBar.DotRadiusY;
      pBar.OnDotRadiusYChanged(oldDotRadiusY, newDotRadiusY);
    }

    /// <summary>
    /// Handles changes to the DotWidth property.
    /// </summary>
    /// <param name="d">
    /// FluidProgressBar
    /// </param>
    /// <param name="e">
    /// DependencyProperty changed event arguments
    /// </param>
    private static void OnDotWidthChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
      FluidProgressBar pBar = (FluidProgressBar)d;
      double oldDotWidth = (double)e.OldValue;
      double newDotWidth = pBar.DotWidth;
      pBar.OnDotWidthChanged(oldDotWidth, newDotWidth);
    }

    /// <summary>
    /// Handles changes to the DurationA property.
    /// </summary>
    /// <param name="d">
    /// FluidProgressBar
    /// </param>
    /// <param name="e">
    /// DependencyProperty changed event arguments
    /// </param>
    private static void OnDurationAChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
      FluidProgressBar pBar = (FluidProgressBar)d;
      Duration oldDurationA = (Duration)e.OldValue;
      Duration newDurationA = pBar.DurationA;
      pBar.OnDurationAChanged(oldDurationA, newDurationA);
    }

    /// <summary>
    /// Handles changes to the DurationB property.
    /// </summary>
    /// <param name="d">
    /// FluidProgressBar
    /// </param>
    /// <param name="e">
    /// DependencyProperty changed event arguments
    /// </param>
    private static void OnDurationBChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
      FluidProgressBar pBar = (FluidProgressBar)d;
      Duration oldDurationB = (Duration)e.OldValue;
      Duration newDurationB = pBar.DurationB;
      pBar.OnDurationBChanged(oldDurationB, newDurationB);
    }

    /// <summary>
    /// Handles changes to the DurationC property.
    /// </summary>
    /// <param name="d">
    /// FluidProgressBar
    /// </param>
    /// <param name="e">
    /// DependencyProperty changed event arguments
    /// </param>
    private static void OnDurationCChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
      FluidProgressBar pBar = (FluidProgressBar)d;
      Duration oldDurationC = (Duration)e.OldValue;
      Duration newDurationC = pBar.DurationC;
      pBar.OnDurationCChanged(oldDurationC, newDurationC);
    }

    /// <summary>
    /// Handles changes to the KeyFrameA property.
    /// </summary>
    /// <param name="d">
    /// FluidProgressBar
    /// </param>
    /// <param name="e">
    /// DependencyProperty changed event arguments
    /// </param>
    private static void OnKeyFrameAChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
      FluidProgressBar pBar = (FluidProgressBar)d;
      double oldKeyFrameA = (double)e.OldValue;
      double newKeyFrameA = pBar.KeyFrameA;
      pBar.OnKeyFrameAChanged(oldKeyFrameA, newKeyFrameA);
    }

    /// <summary>
    /// Handles changes to the KeyFrameB property.
    /// </summary>
    /// <param name="d">
    /// FluidProgressBar
    /// </param>
    /// <param name="e">
    /// DependencyProperty changed event arguments
    /// </param>
    private static void OnKeyFrameBChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
      FluidProgressBar pBar = (FluidProgressBar)d;
      double oldKeyFrameB = (double)e.OldValue;
      double newKeyFrameB = pBar.KeyFrameB;
      pBar.OnKeyFrameBChanged(oldKeyFrameB, newKeyFrameB);
    }

    /// <summary>
    /// Handles changes to the Oscillate property.
    /// </summary>
    /// <param name="d">
    /// FluidProgressBar
    /// </param>
    /// <param name="e">
    /// DependencyProperty changed event arguments
    /// </param>
    private static void OnOscillateChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
      FluidProgressBar pBar = (FluidProgressBar)d;
      bool oldOscillate = (bool)e.OldValue;
      bool newOscillate = pBar.Oscillate;
      pBar.OnOscillateChanged(oldOscillate, newOscillate);
    }

    /// <summary>
    /// Handles changes to the ReverseDuration property.
    /// </summary>
    /// <param name="d">
    /// FluidProgressBar
    /// </param>
    /// <param name="e">
    /// DependencyProperty changed event arguments
    /// </param>
    private static void OnReverseDurationChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
      FluidProgressBar pBar = (FluidProgressBar)d;
      Duration oldReverseDuration = (Duration)e.OldValue;
      Duration newReverseDuration = pBar.ReverseDuration;
      pBar.OnReverseDurationChanged(oldReverseDuration, newReverseDuration);
    }

    /// <summary>
    /// Handles changes to the TotalDuration property.
    /// </summary>
    /// <param name="d">
    /// FluidProgressBar
    /// </param>
    /// <param name="e">
    /// DependencyProperty changed event arguments
    /// </param>
    private static void OnTotalDurationChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
      FluidProgressBar pBar = (FluidProgressBar)d;
      Duration oldTotalDuration = (Duration)e.OldValue;
      Duration newTotalDuration = pBar.TotalDuration;
      pBar.OnTotalDurationChanged(oldTotalDuration, newTotalDuration);
    }

    /// <summary>
    ///   Obtains the keyframes for each animation in the storyboard so that
    ///   they can be updated when required.
    /// </summary>
    private void GetKeyFramesFromStoryboard()
    {
      this.sb = (Storyboard)this.Resources["FluidStoryboard"];
      if (this.sb != null)
      {
        foreach (Timeline timeline in this.sb.Children)
        {
          DoubleAnimationUsingKeyFrames dakeys = timeline as DoubleAnimationUsingKeyFrames;
          if (dakeys != null)
          {
            string targetName = Storyboard.GetTargetName(dakeys);
            this.ProcessDoubleAnimationWithKeys(dakeys, !targetName.StartsWith("Trans"));
          }
        }
      }
    }

    /// <summary>
    /// Handles the IsVisibleChanged event
    /// </summary>
    /// <param name="sender">
    /// Sender
    /// </param>
    /// <param name="e">
    /// EventArgs
    /// </param>
    private void OnIsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
    {
      if (this.Visibility == Visibility.Visible)
      {
        this.UpdateKeyFrames();
        this.StartFluidAnimation();
      }
      else
      {
        this.StopFluidAnimation();
      }
    }

    /// <summary>
    /// Handles the Loaded event
    /// </summary>
    /// <param name="sender">
    /// Sender
    /// </param>
    /// <param name="e">
    /// EventArgs
    /// </param>
    private void OnLoaded(object sender, RoutedEventArgs e)
    {
      // Update the key frames
      this.UpdateKeyFrames();

      // Start the animation
      this.StartFluidAnimation();
    }

    /// <summary>
    /// Handles the SizeChanged event
    /// </summary>
    /// <param name="sender">
    /// Sender
    /// </param>
    /// <param name="e">
    /// EventArgs
    /// </param>
    private void OnSizeChanged(object sender, SizeChangedEventArgs e)
    {
      // Restart the animation
      this.RestartStoryboardAnimation();
    }

    /// <summary>
    /// Gets the keyframes in the given animation and stores them in a map
    /// </summary>
    /// <param name="dakeys">
    /// Animation containg keyframes
    /// </param>
    /// <param name="isOpacityAnim">
    /// Flag to indicate whether the animation targets the opacity or the translate transform
    /// </param>
    private void ProcessDoubleAnimationWithKeys(DoubleAnimationUsingKeyFrames dakeys, bool isOpacityAnim = false)
    {
      // Get all the keyframes in the instance.
      for (int i = 0; i < dakeys.KeyFrames.Count; i++)
      {
        DoubleKeyFrame frame = dakeys.KeyFrames[i];

        Dictionary<int, KeyFrameDetails> targetMap = null;

        if (isOpacityAnim)
        {
          targetMap = this.opKeyFrameMap;
        }
        else
        {
          targetMap = this.keyFrameMap;
        }

        if (!targetMap.ContainsKey(i))
        {
          targetMap[i] = new KeyFrameDetails() { KeyFrames = new List<DoubleKeyFrame>() };
        }

        // Update the keyframe time and add it to the map
        targetMap[i].KeyFrameTime = frame.KeyTime;
        targetMap[i].KeyFrames.Add(frame);
      }
    }

    /// <summary>
    ///   Stops the animation, updates the keyframes and starts the animation
    /// </summary>
    private void RestartStoryboardAnimation()
    {
      this.StopFluidAnimation();
      this.UpdateKeyFrames();
      this.StartFluidAnimation();
    }

    /// <summary>
    ///   Starts the animation
    /// </summary>
    private void StartFluidAnimation()
    {
      if ((this.sb != null) && (!this.isStoryboardRunning))
      {
        this.sb.Begin();
        this.isStoryboardRunning = true;
      }
    }

    /// <summary>
    ///   Stops the animation
    /// </summary>
    private void StopFluidAnimation()
    {
      if ((this.sb != null) && this.isStoryboardRunning)
      {
        // Move the timeline to the end and stop the animation
        this.sb.SeekAlignedToLastTick(TimeSpan.FromSeconds(0));
        this.sb.Stop();
        this.isStoryboardRunning = false;
      }
    }

    /// <summary>
    /// Update the key value of the keyframes stored in the map
    /// </summary>
    /// <param name="key">
    /// Key of the dictionary
    /// </param>
    /// <param name="newValue">
    /// New value to be given to the key value of the keyframes
    /// </param>
    private void UpdateKeyFrame(int key, double newValue)
    {
      if (this.keyFrameMap.ContainsKey(key))
      {
        foreach (var frame in this.keyFrameMap[key].KeyFrames)
        {
          if (frame is LinearDoubleKeyFrame)
          {
            frame.SetValue(LinearDoubleKeyFrame.ValueProperty, newValue);
          }
          else if (frame is EasingDoubleKeyFrame)
          {
            frame.SetValue(EasingDoubleKeyFrame.ValueProperty, newValue);
          }
        }
      }
    }

    /// <summary>
    ///   Update the key value of each keyframe based on the current width of the FluidProgressBar
    /// </summary>
    private void UpdateKeyFrames()
    {
      // Get the current width of the FluidProgressBar
      double width = this.ActualWidth;

      // Update the values only if the current width is greater than Zero and is visible
      if ((width > 0.0) && (this.Visibility == Visibility.Visible))
      {
        double Point0 = -10;
        double PointA = width * this.KeyFrameA;
        double PointB = width * this.KeyFrameB;
        double PointC = width + 10;

        // Update the keyframes stored in the map
        this.UpdateKeyFrame(0, Point0);
        this.UpdateKeyFrame(1, PointA);
        this.UpdateKeyFrame(2, PointB);
        this.UpdateKeyFrame(3, PointC);
      }
    }

    /// <summary>
    /// Updates the duration of each of the keyframes stored in the map
    /// </summary>
    /// <param name="key">
    /// Key of the dictionary
    /// </param>
    /// <param name="newDuration">
    /// New value to be given to the duration value of the keyframes
    /// </param>
    private void UpdateKeyTime(int key, Duration newDuration)
    {
      if (this.keyFrameMap.ContainsKey(key))
      {
        KeyTime newKeyTime = KeyTime.FromTimeSpan(newDuration.TimeSpan);
        this.keyFrameMap[key].KeyFrameTime = newKeyTime;

        foreach (var frame in this.keyFrameMap[key].KeyFrames)
        {
          if (frame is LinearDoubleKeyFrame)
          {
            frame.SetValue(LinearDoubleKeyFrame.KeyTimeProperty, newKeyTime);
          }
          else if (frame is EasingDoubleKeyFrame)
          {
            frame.SetValue(EasingDoubleKeyFrame.KeyTimeProperty, newKeyTime);
          }
        }
      }
    }

    /// <summary>
    /// Updates the duration of each of the keyframes stored in the map
    /// </summary>
    /// <param name="key">
    /// Key of the dictionary
    /// </param>
    /// <param name="newDuration">
    /// New value to be given to the duration value of the keyframes
    /// </param>
    private void UpdateKeyTimes(int key, Duration newDuration)
    {
      switch (key)
      {
        case 1:
          this.UpdateKeyTime(1, newDuration);
          this.UpdateKeyTime(2, newDuration + this.DurationB);
          this.UpdateKeyTime(3, newDuration + this.DurationB + this.DurationC);
          break;

        case 2:
          this.UpdateKeyTime(2, this.DurationA + newDuration);
          this.UpdateKeyTime(3, this.DurationA + newDuration + this.DurationC);
          break;

        case 3:
          this.UpdateKeyTime(3, this.DurationA + this.DurationB + newDuration);
          break;
      }

      // Update the opacity animation duration based on the complete duration
      // of the animation
      this.UpdateOpacityKeyTime(1, this.DurationA + this.DurationB + this.DurationC);
    }

    /// <summary>
    /// Updates the duration of the second keyframe of all the opacity animations
    /// </summary>
    /// <param name="key">
    /// Key of the dictionary
    /// </param>
    /// <param name="newDuration">
    /// New value to be given to the duration value of the keyframes
    /// </param>
    private void UpdateOpacityKeyTime(int key, Duration newDuration)
    {
      if (!this.opKeyFrameMap.ContainsKey(key))
      {
        return;
      }

      KeyTime newKeyTime = KeyTime.FromTimeSpan(newDuration.TimeSpan);
      this.opKeyFrameMap[key].KeyFrameTime = newKeyTime;

      foreach (var frame in this.opKeyFrameMap[key].KeyFrames)
      {
        if (frame is DiscreteDoubleKeyFrame)
        {
          frame.SetValue(DiscreteDoubleKeyFrame.KeyTimeProperty, newKeyTime);
        }
      }
    }

    /// <summary>
    /// Updates the delay between consecutive timelines
    /// </summary>
    /// <param name="newDelay">
    /// Delay duration
    /// </param>
    private void UpdateTimelineDelay(Duration newDelay)
    {
      Duration nextDelay = new Duration(TimeSpan.FromSeconds(0));

      if (this.sb == null)
      {
        return;
      }

      for (int i = 0; i < this.sb.Children.Count; i++)
      {
        // The first five animations are for translation
        // The next five animations are for opacity
        if (i == 5)
        {
          nextDelay = newDelay;
        }
        else
        {
          nextDelay += newDelay;
        }

        var timeline = this.sb.Children[i] as DoubleAnimationUsingKeyFrames;

        if (timeline != null)
        {
          timeline.SetValue(DoubleAnimationUsingKeyFrames.BeginTimeProperty, nextDelay.TimeSpan);
        }
      }
    }

    #endregion

    /// <summary>
    /// </summary>
    private class KeyFrameDetails
    {
      #region Public Properties

      /// <summary>
      ///   Gets or sets the key frame time.
      /// </summary>
      /// <value>
      ///   The key frame time.
      /// </value>
      public KeyTime KeyFrameTime { get; set; }

      /// <summary>
      ///   Gets or sets the key frames.
      /// </summary>
      /// <value>
      ///   The key frames.
      /// </value>
      public List<DoubleKeyFrame> KeyFrames { get; set; }

      #endregion
    }
  }
}