using Microsoft.Kinect;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using PPt = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using System.Threading;


namespace KinectHandTracking
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        PPt.Application pptApplication;
        // Define Presentation object
        PPt.Presentation presentation;
        // Define Slide collection
        PPt.Slides slides;
        PPt.Slide slide;

        // Slide count
        int slidescount;
        // slide index
        int slideIndex;
        string leftHandState = "-";
        //bool check;

        // DateTime getTimeDelay = new DateTime();

        #region Members

        KinectSensor _sensor;
        MultiSourceFrameReader _reader;
        IList<Body> _bodies;

        #endregion

        #region Constructor

        
        /*void checkCheck()
        {
            check = true;
            Thread.Sleep(1000);
        }
        */

        void checkPP()
        {
            try
            {
                // Get Running PowerPoint Application object
                pptApplication = Marshal.GetActiveObject("PowerPoint.Application") as PPt.Application;
            }
            catch
            {
                tblLeftHandState.Text = "Please Run PowerPoint first";
            }
            if (pptApplication != null)
            {
                // Get Presentation Object
                presentation = pptApplication.ActivePresentation;
                // Get Slide collection object
                slides = presentation.Slides;
                // Get Slide count
                slidescount = slides.Count;
                // Get current selected slide 
                try
                {
                    // Get selected slide object in normal view
                    slide = slides[pptApplication.ActiveWindow.Selection.SlideRange.SlideNumber];
                }
                catch
                {
                    // Get selected slide object in reading view
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                }
            }

        }

        public MainWindow()
        {
            InitializeComponent();
        }

        #endregion

        #region Event handlers

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            _sensor = KinectSensor.GetDefault();
            
           

            if (_sensor != null)
            {
                _sensor.Open();

                _reader = _sensor.OpenMultiSourceFrameReader(FrameSourceTypes.Color | FrameSourceTypes.Depth | FrameSourceTypes.Infrared | FrameSourceTypes.Body);
                _reader.MultiSourceFrameArrived += Reader_MultiSourceFrameArrived;
            }
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            if (_reader != null)
            {
                _reader.Dispose();
            }

            if (_sensor != null)
            {
                _sensor.Close();
            }
        }





        void Reader_MultiSourceFrameArrived(object sender, MultiSourceFrameArrivedEventArgs e)
        {

            /*getTimeDelay = DateTime.Now;
            Thread newThread = new Thread(checkCheck);
            newThread.Start();
            */

            var reference = e.FrameReference.AcquireFrame();

            // Color
            using (var frame = reference.ColorFrameReference.AcquireFrame())
            {
                if (frame != null)
                {
                    camera.Source = frame.ToBitmap();
                }
            }

            
            // Body
            using (var frame = reference.BodyFrameReference.AcquireFrame())
            {
                if (frame != null)
                {
                    canvas.Children.Clear();

                    _bodies = new Body[frame.BodyFrameSource.BodyCount];

                    frame.GetAndRefreshBodyData(_bodies);

                    foreach (var body in _bodies)
                    {
                        if (body != null)
                        {
                            if (body.IsTracked)
                            {
                                // Find the joints
                                Joint handRight = body.Joints[JointType.HandRight];
                                Joint thumbRight = body.Joints[JointType.ThumbRight];

                                Joint handLeft = body.Joints[JointType.HandLeft];
                                Joint thumbLeft = body.Joints[JointType.ThumbLeft];

                                // Draw hands and thumbs
                                canvas.DrawHand(handRight, _sensor.CoordinateMapper);
                                canvas.DrawHand(handLeft, _sensor.CoordinateMapper);
                                canvas.DrawThumb(thumbRight, _sensor.CoordinateMapper);
                                canvas.DrawThumb(thumbLeft, _sensor.CoordinateMapper);

                                // Find the hand states
                                
                                    // Previous slide
                                    if (body.Joints[JointType.HandRight].Position.X - body.Joints[JointType.ShoulderLeft].Position.X > 0 &&
                                        body.Joints[JointType.HandRight].Position.X - body.Joints[JointType.Head].Position.X < 0 &&
                                        body.Joints[JointType.HandRight].Position.Y - body.Joints[JointType.SpineShoulder].Position.Y > 0 &&
                                        body.Joints[JointType.HandRight].Position.Y - body.Joints[JointType.Head].Position.Y < 0 &&
                                        body.HandRightState == HandState.Open
                                        )
                                    {
                                        checkPP();
                                        tblLeftHandState.Text = "Previous animation";
                                        slideIndex = slide.SlideIndex - 1;
                                        if (slideIndex >= 1)
                                        {
                                            try
                                            {
                                                slide = slides[slideIndex];
                                                slides[slideIndex].Select();
                                            }
                                            catch
                                            {
                                                pptApplication.SlideShowWindows[1].View.Previous();
                                                slide = pptApplication.SlideShowWindows[1].View.Slide;
                                            }
                                       
                                    }
                                    
                                        Thread.Sleep(500);
                                    }

                                    // Next slide
                                    else if (body.Joints[JointType.HandLeft].Position.X - body.Joints[JointType.Head].Position.X>0 &&
                                             body.Joints[JointType.HandLeft].Position.X - body.Joints[JointType.ShoulderRight].Position.X < 0 &&
                                             body.Joints[JointType.HandLeft].Position.Y - body.Joints[JointType.SpineShoulder].Position.Y > 0 &&
                                             body.Joints[JointType.HandLeft].Position.Y - body.Joints[JointType.Head].Position.Y < 0 &&
                                             body.HandLeftState == HandState.Open)
                                    {
                                        checkPP();
                                        tblLeftHandState.Text = "Next animation";
                                        slideIndex = slide.SlideIndex + 1;
                                        if (slideIndex <= slidescount)
                                        {
                                            try
                                            {
                                                slide = slides[slideIndex];
                                                slides[slideIndex].Select();
                                            }
                                            catch
                                            {
                                                pptApplication.SlideShowWindows[1].View.Next();
                                                slide = pptApplication.SlideShowWindows[1].View.Slide;
                                            }
                                        
                                        }
                                                                          Thread.Sleep(500);
                                    }

                                    // First slide
                                    else if (body.HandLeftState == HandState.Lasso &&
                                        body.HandRightState == HandState.Lasso)
                                    {
                                        checkPP();
                                      
                                        tblLeftHandState.Text = "Frist Slide";
                                        try
                                        {
                                            // Call Select method to select first slide in normal view
                                            slides[1].Select();
                                            slide = slides[1];
                                        }
                                        catch
                                        {
                                            // Transform to first page in reading view
                                            pptApplication.SlideShowWindows[1].View.First();
                                            slide = pptApplication.SlideShowWindows[1].View.Slide;
                                        }
                                          Thread.Sleep(500);
                                    }

                                    // Last slide
                                    else if (body.HandLeftState == HandState.Closed &&
                                        body.HandRightState == HandState.Lasso)
                                    {
                                        checkPP();
                                        tblLeftHandState.Text = "Last slide";
                                        slideIndex = slide.SlideIndex + 1;
                                        try
                                        {
                                            slides[slidescount].Select();
                                            slide = slides[slidescount];
                                        }
                                        catch
                                        {
                                            pptApplication.SlideShowWindows[1].View.Last();
                                            slide = pptApplication.SlideShowWindows[1].View.Slide;
                                        }
                                         Thread.Sleep(1000);
                                    }
                                    //check = false;

                                }
   
                            
                        }
                    }
                }
            }
        }

        #endregion
    }
}
