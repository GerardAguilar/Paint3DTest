//******************************************************************************
//
// Copyright (c) 2018 Microsoft Corporation. All rights reserved.
//
// This code is licensed under the MIT License (MIT).
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//
//******************************************************************************

using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Appium.Interactions;
using OpenQA.Selenium.Appium.Windows;

// Define an alias to OpenQA.Selenium.Appium.Interactions.PointerInputDevice to hide
// inherited OpenQA.Selenium.Interactions.PointerInputDevice that causes ambiguity.
// In the future, all functions of OpenQA.Selenium.Appium.Interactions should be moved
// up to OpenQA.Selenium.Interactions and this alias can simply be removed.
using PointerInputDevice = OpenQA.Selenium.Appium.Interactions.PointerInputDevice;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace Paint3DTest
{
    [TestClass]
    public class ScenarioZoom : Paint3DSession
    {
        private WindowsElement zoomInteractor;
        private WindowsElement zoomScaleTextBox;

        private static class User32
        {
            [DllImport("User32.dll")]
            internal static extern IntPtr SetForegroundWindow(IntPtr hWnd);

            [DllImport("user32.dll")]
            internal static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

            internal static readonly IntPtr InvalidHandleValue = IntPtr.Zero;
            internal const int SW_MAXIMIZE = 3;
        }
        public void Activate(int processId)
        {
            //Process currentProcess = Process.GetCurrentProcess();
            Process currentProcess = Process.GetProcessById(processId);
            IntPtr hWnd = currentProcess.MainWindowHandle;
            if (hWnd != User32.InvalidHandleValue)
            {
                User32.SetForegroundWindow(hWnd);
                //User32.ShowWindow(hWnd, User32.SW_MAXIMIZE);
            }
            else {
                Console.WriteLine("hWnd is invalid");
            }
        }

        [TestMethod]
        public void ZoomingInMultiTouch()
        {
            // Drag a touch contact diagonally in NE direction distancing apart from the other contact point
            PointerInputDevice touch1 = new PointerInputDevice(PointerKind.Touch);
            ActionSequence touch1Sequence = new ActionSequence(touch1, 0);
            touch1Sequence.AddAction(touch1.CreatePointerMove(zoomInteractor, 50, -50, TimeSpan.Zero));
            touch1Sequence.AddAction(touch1.CreatePointerDown(PointerButton.TouchContact));
            touch1Sequence.AddAction(touch1.CreatePointerMove(zoomInteractor, 55, -55, TimeSpan.Zero));
            touch1Sequence.AddAction(touch1.CreatePointerMove(zoomInteractor, 60, -60, TimeSpan.Zero));
            touch1Sequence.AddAction(touch1.CreatePointerMove(zoomInteractor, 65, -65, TimeSpan.Zero));
            touch1Sequence.AddAction(touch1.CreatePointerMove(zoomInteractor, 70, -70, TimeSpan.Zero));
            touch1Sequence.AddAction(touch1.CreatePointerMove(zoomInteractor, 75, -75, TimeSpan.Zero));
            touch1Sequence.AddAction(touch1.CreatePointerMove(zoomInteractor, 80, -80, TimeSpan.Zero));
            touch1Sequence.AddAction(touch1.CreatePointerUp(PointerButton.TouchContact));

            // Drag a touch contact diagonally in SW direction distancing apart from the other contact point
            PointerInputDevice touch2 = new PointerInputDevice(PointerKind.Touch);
            ActionSequence touch2Sequence = new ActionSequence(touch2, 0);
            touch2Sequence.AddAction(touch2.CreatePointerMove(zoomInteractor, -50, 50, TimeSpan.Zero));
            touch2Sequence.AddAction(touch2.CreatePointerDown(PointerButton.TouchContact));
            touch2Sequence.AddAction(touch2.CreatePointerMove(zoomInteractor, -55, 55, TimeSpan.Zero));
            touch2Sequence.AddAction(touch2.CreatePointerMove(zoomInteractor, -60, 60, TimeSpan.Zero));
            touch2Sequence.AddAction(touch2.CreatePointerMove(zoomInteractor, -65, 65, TimeSpan.Zero));
            touch2Sequence.AddAction(touch2.CreatePointerMove(zoomInteractor, -70, 70, TimeSpan.Zero));
            touch2Sequence.AddAction(touch2.CreatePointerMove(zoomInteractor, -75, 75, TimeSpan.Zero));
            touch2Sequence.AddAction(touch2.CreatePointerMove(zoomInteractor, -80, 80, TimeSpan.Zero));
            touch2Sequence.AddAction(touch2.CreatePointerUp(PointerButton.TouchContact));

            // Perform the 2 fingers zoom in (expand) multi-touch sequences defined above
            session.PerformActions(new List<ActionSequence> { touch1Sequence, touch2Sequence });

            // Ensure that the zoom level now is greater than 100%
            Assert.IsTrue(int.Parse(zoomScaleTextBox.Text) > 100);
        }

        [TestMethod]
        public void ZoomingInMultiTouchWithInterpolation()
        {
            // Set pointer move Duration to 300 ms to implicitly generate 6 interpolation moves that are performed every 50 ms
            TimeSpan moveDuration = TimeSpan.FromMilliseconds(300);

            // Drag a touch contact diagonally in NE direction distancing apart from the other contact point
            PointerInputDevice touch1 = new PointerInputDevice(PointerKind.Touch);
            ActionSequence touch1Sequence = new ActionSequence(touch1, 0);
            touch1Sequence.AddAction(touch1.CreatePointerMove(zoomInteractor, 50, -50, TimeSpan.Zero));
            touch1Sequence.AddAction(touch1.CreatePointerDown(PointerButton.TouchContact));
            touch1Sequence.AddAction(touch1.CreatePointerMove(zoomInteractor, 80, -80, moveDuration));
            touch1Sequence.AddAction(touch1.CreatePointerUp(PointerButton.TouchContact));

            // Drag a touch contact diagonally in SW direction distancing apart from the other contact point
            PointerInputDevice touch2 = new PointerInputDevice(PointerKind.Touch);
            ActionSequence touch2Sequence = new ActionSequence(touch2, 0);
            touch2Sequence.AddAction(touch2.CreatePointerMove(zoomInteractor, -50, 50, TimeSpan.Zero));
            touch2Sequence.AddAction(touch2.CreatePointerDown(PointerButton.TouchContact));
            touch2Sequence.AddAction(touch2.CreatePointerMove(zoomInteractor, -80, 80, moveDuration));
            touch2Sequence.AddAction(touch2.CreatePointerUp(PointerButton.TouchContact));

            // Perform the 2 fingers zoom in (expand) multi-touch sequences defined above
            session.PerformActions(new List<ActionSequence> { touch1Sequence, touch2Sequence });

            // Ensure that the zoom level now is greater than 100%
            Assert.IsTrue(int.Parse(zoomScaleTextBox.Text) > 100);
        }

        [TestMethod]
        public void ZoomingOutMultiTouch()
        {
            // Drag a touch contact diagonally in SW direction approaching the other contact point
            PointerInputDevice touch1 = new PointerInputDevice(PointerKind.Touch);
            ActionSequence touch1Sequence = new ActionSequence(touch1, 0);
            touch1Sequence.AddAction(touch1.CreatePointerMove(zoomInteractor, 50, -50, TimeSpan.Zero));
            touch1Sequence.AddAction(touch1.CreatePointerDown(PointerButton.TouchContact));
            touch1Sequence.AddAction(touch1.CreatePointerMove(zoomInteractor, 45, -45, TimeSpan.Zero));
            touch1Sequence.AddAction(touch1.CreatePointerMove(zoomInteractor, 40, -40, TimeSpan.Zero));
            touch1Sequence.AddAction(touch1.CreatePointerMove(zoomInteractor, 35, -35, TimeSpan.Zero));
            touch1Sequence.AddAction(touch1.CreatePointerMove(zoomInteractor, 30, -30, TimeSpan.Zero));
            touch1Sequence.AddAction(touch1.CreatePointerMove(zoomInteractor, 25, -25, TimeSpan.Zero));
            touch1Sequence.AddAction(touch1.CreatePointerMove(zoomInteractor, 20, -20, TimeSpan.Zero));
            touch1Sequence.AddAction(touch1.CreatePointerUp(PointerButton.TouchContact));

            // Drag a touch contact diagonally in NE direction approaching the other contact point
            PointerInputDevice touch2 = new PointerInputDevice(PointerKind.Touch);
            ActionSequence touch2Sequence = new ActionSequence(touch2, 0);
            touch2Sequence.AddAction(touch2.CreatePointerMove(zoomInteractor, -50, 50, TimeSpan.Zero));
            touch2Sequence.AddAction(touch2.CreatePointerDown(PointerButton.TouchContact));
            touch2Sequence.AddAction(touch2.CreatePointerMove(zoomInteractor, -45, 45, TimeSpan.Zero));
            touch2Sequence.AddAction(touch2.CreatePointerMove(zoomInteractor, -40, 40, TimeSpan.Zero));
            touch2Sequence.AddAction(touch2.CreatePointerMove(zoomInteractor, -35, 35, TimeSpan.Zero));
            touch2Sequence.AddAction(touch2.CreatePointerMove(zoomInteractor, -30, 30, TimeSpan.Zero));
            touch2Sequence.AddAction(touch2.CreatePointerMove(zoomInteractor, -25, 25, TimeSpan.Zero));
            touch2Sequence.AddAction(touch2.CreatePointerMove(zoomInteractor, -20, 20, TimeSpan.Zero));
            touch2Sequence.AddAction(touch2.CreatePointerUp(PointerButton.TouchContact));

            // Perform the 2 fingers zoom out (pinch) multi-touch sequences defined above
            session.PerformActions(new List<ActionSequence> { touch1Sequence, touch2Sequence });

            // Ensure that the zoom level now is less than 100%
            Assert.IsTrue(int.Parse(zoomScaleTextBox.Text) < 100);
        }

        [TestMethod]
        public void ZoomingOutMultiTouchWithInterpolation()
        {
            // Set pointer move Duration to 300 ms to implicitly generate 6 interpolation moves that are performed every 50 ms
            TimeSpan moveDuration = TimeSpan.FromMilliseconds(300);

            // Drag a touch contact diagonally in SW direction approaching the other contact point
            PointerInputDevice touch1 = new PointerInputDevice(PointerKind.Touch);
            ActionSequence touch1Sequence = new ActionSequence(touch1, 0);
            touch1Sequence.AddAction(touch1.CreatePointerMove(zoomInteractor, 50, -50, TimeSpan.Zero));
            touch1Sequence.AddAction(touch1.CreatePointerDown(PointerButton.TouchContact));
            touch1Sequence.AddAction(touch1.CreatePointerMove(zoomInteractor, 20, -20, moveDuration));
            touch1Sequence.AddAction(touch1.CreatePointerUp(PointerButton.TouchContact));

            // Drag a touch contact diagonally in NE direction approaching the other contact point
            PointerInputDevice touch2 = new PointerInputDevice(PointerKind.Touch);
            ActionSequence touch2Sequence = new ActionSequence(touch2, 0);
            touch2Sequence.AddAction(touch2.CreatePointerMove(zoomInteractor, -50, 50, TimeSpan.Zero));
            touch2Sequence.AddAction(touch2.CreatePointerDown(PointerButton.TouchContact));
            touch2Sequence.AddAction(touch2.CreatePointerMove(zoomInteractor, -20, 20, moveDuration));
            touch2Sequence.AddAction(touch2.CreatePointerUp(PointerButton.TouchContact));

            // Perform the 2 fingers zoom out (pinch) multi-touch sequences defined above
            session.PerformActions(new List<ActionSequence> { touch1Sequence, touch2Sequence });

            // Ensure that the zoom level now is less than 100%
            Assert.IsTrue(int.Parse(zoomScaleTextBox.Text) < 100);
        }

        [TestMethod]
        public void TapTest()
        {
            PointerInputDevice touch = new PointerInputDevice(PointerKind.Touch);
            ActionSequence touchSequence = new ActionSequence(touch, 0);
            touchSequence.AddAction(touch.CreatePointerMove(zoomInteractor, 0, 0, TimeSpan.Zero));
            touchSequence.AddAction(touch.CreatePointerDown(PointerButton.TouchContact));
            touchSequence.AddAction(touch.CreatePointerUp(PointerButton.TouchContact));
            touchSequence.AddAction(touch.CreatePointerDown(PointerButton.TouchContact));
            touchSequence.AddAction(touch.CreatePointerUp(PointerButton.TouchContact));
            session.PerformActions(new List<ActionSequence> { touchSequence });

            Assert.IsTrue(true);
        }

        [TestMethod]
        public void TapTest1()
        {
            PointerInputDevice touch = new PointerInputDevice(PointerKind.Touch);
            ActionSequence touchSequence = new ActionSequence(touch, 0);
            touchSequence.AddAction(touch.CreatePointerMove(zoomInteractor, 0, 0, TimeSpan.Zero));
            touchSequence.AddAction(touch.CreatePointerDown(PointerButton.TouchContact));
            touchSequence.AddAction(touch.CreatePointerUp(PointerButton.TouchContact));
            touchSequence.AddAction(touch.CreatePointerMove(zoomInteractor, 10, 0, TimeSpan.Zero));
            touchSequence.AddAction(touch.CreatePointerDown(PointerButton.TouchContact));
            touchSequence.AddAction(touch.CreatePointerUp(PointerButton.TouchContact));
            touchSequence.AddAction(touch.CreatePointerMove(zoomInteractor, 20, 0, TimeSpan.Zero));
            touchSequence.AddAction(touch.CreatePointerDown(PointerButton.TouchContact));
            touchSequence.AddAction(touch.CreatePointerUp(PointerButton.TouchContact));
            touchSequence.AddAction(touch.CreatePointerMove(zoomInteractor, 30, 0, TimeSpan.Zero));
            touchSequence.AddAction(touch.CreatePointerDown(PointerButton.TouchContact));
            touchSequence.AddAction(touch.CreatePointerUp(PointerButton.TouchContact));
            session.PerformActions(new List<ActionSequence> { touchSequence });
            Assert.IsTrue(true);
        }

        [TestMethod]
        public void TapTest3() {
            Tap(40, 0);
            Tap(50, 10);
            Tap(70, 20);
            Tap(100, 30);
            Assert.IsTrue(true);            
        }

        [TestMethod]
        public void TapTest4() {
            Activate(21572);
            TapTest3();
        }

        [TestMethod]
        public void TapTest2()
        {
            //process ID
            Activate(3316);

            // Draw a circle with radius 300 and 40 (x, y) points
            const int radius = 300;
            const int points = 40;

            // Draw the circle with a single touch actions
            PointerInputDevice touchContact = new PointerInputDevice(PointerKind.Touch);
            ActionSequence touchSequence = new ActionSequence(touchContact, 0);
            //touchSequence.AddAction(touchContact.CreatePointerMove(inkCanvas, 0, -radius, TimeSpan.Zero));
            //touchSequence.AddAction(touchContact.CreatePointerMove(paintCanvas, 0, -radius, TimeSpan.Zero));
            //touchSequence.AddAction(touchContact.CreatePointerMove(CoordinateOrigin.Viewport, 0, 0, TimeSpan.Zero));
            //touchSequence.AddAction(touchContact.CreatePointerDown(PointerButton.TouchContact));

            //touchSequence.AddAction(touchContact.CreatePointerDown(PointerButton.TouchContact));
            //for (double angle = 0; angle <= 2 * Math.PI; angle += 2 * Math.PI / points)
            //{
            //    //touchSequence.AddAction(touchContact.CreatePointerMove(inkCanvas, (int)(Math.Sin(angle) * radius), -(int)(Math.Cos(angle) * radius), TimeSpan.Zero));

            //    touchSequence.AddAction(touchContact.CreatePointerMove(CoordinateOrigin.Viewport, (int)(Math.Sin(angle) * radius), -(int)(Math.Cos(angle) * radius), TimeSpan.Zero));

            //    //touchSequence.AddAction(touchContact.CreatePointerMove(paintCanvas, 0, -radius, TimeSpan.Zero));
            //}
            //touchSequence.AddAction(touchContact.CreatePointerUp(PointerButton.TouchContact));

            touchSequence.AddAction(touchContact.CreatePointerDown(PointerButton.TouchContact));
            touchSequence.AddAction(touchContact.CreatePointerMove(CoordinateOrigin.Viewport, 1, 1, TimeSpan.Zero));
            touchSequence.AddAction(touchContact.CreatePointerUp(PointerButton.TouchContact));
            touchSequence.AddAction(touchContact.CreatePointerDown(PointerButton.TouchContact));
            touchSequence.AddAction(touchContact.CreatePointerMove(CoordinateOrigin.Viewport, 400, 400, TimeSpan.Zero));
            touchSequence.AddAction(touchContact.CreatePointerUp(PointerButton.TouchContact));
            //touchSequence.AddAction(touchContact.CreatePointerDown(PointerButton.TouchContact));
            //touchSequence.AddAction(touchContact.CreatePointerMove(CoordinateOrigin.Viewport, 400, 400, TimeSpan.Zero));
            //touchSequence.AddAction(touchContact.CreatePointerUp(PointerButton.TouchContact));
            //touchSequence.AddAction(touchContact.CreatePointerDown(PointerButton.TouchContact));
            //touchSequence.AddAction(touchContact.CreatePointerMove(CoordinateOrigin.Viewport, 400, 400, TimeSpan.Zero));
            //touchSequence.AddAction(touchContact.CreatePointerUp(PointerButton.TouchContact));


            List<ActionSequence> actionList = new List<ActionSequence> { touchSequence };
            for (int i = 0; i < actionList.Count; i++)
            {
                Console.WriteLine(actionList[i].ToString());
            }
            session.PerformActions(actionList);
        }

        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            // Create session to launch or bring up Paint 3D application
            //Setup(context);
            Setup2(context, "Untitled - Paint 3D");
        }

        [ClassCleanup]
        public static void ClassCleanup()
        {
            //TearDown();
        }

        [TestInitialize]
        public void SetupZoomLevel()
        {
            zoomInteractor = session.FindElementByAccessibilityId("ZoomInteractor");
            //zoomScaleTextBox = session.FindElementByAccessibilityId("ZoomScaleTextBox");

            // Ensure that the zoom level starts at 100%
            //Assert.IsTrue(int.Parse(zoomScaleTextBox.Text) == 100);

            // Draw a circle to help visualize the zoom level changes
            //DrawCircle();
        }

        private void DrawCircle()
        {
            // Draw a circle with radius 300 and 40 (x, y) points
            const int radius = 300;
            const int points = 40;

            // Select the Brushes toolbox to have the Brushes Pane sidebar displayed and ensure that Marker is selected
            session.FindElementByAccessibilityId("Toolbox").FindElementByAccessibilityId("TopBar_ArtTools").Click();
            session.FindElementByAccessibilityId("SidebarWrapper").FindElementByAccessibilityId("Marker3d").Click();

            // Locate the drawing surface
            WindowsElement inkCanvas = session.FindElementByAccessibilityId("InteractorFocusWrapper");
            //WindowsElement paintCanvas = session.FindElementByClassName("Afx:00007FF6908E0000:8");            

            // Draw the circle with a single touch actions
            PointerInputDevice touchContact = new PointerInputDevice(PointerKind.Touch);
            ActionSequence touchSequence = new ActionSequence(touchContact, 0);
            //touchSequence.AddAction(touchContact.CreatePointerMove(inkCanvas, 0, -radius, TimeSpan.Zero));
            //touchSequence.AddAction(touchContact.CreatePointerMove(paintCanvas, 0, -radius, TimeSpan.Zero));
            touchSequence.AddAction(touchContact.CreatePointerMove(CoordinateOrigin.Viewport, 0, -radius, TimeSpan.Zero));
            touchSequence.AddAction(touchContact.CreatePointerDown(PointerButton.TouchContact));
            for (double angle = 0; angle <= 2 * Math.PI; angle += 2 * Math.PI / points)
            {
                //touchSequence.AddAction(touchContact.CreatePointerMove(inkCanvas, (int)(Math.Sin(angle) * radius), -(int)(Math.Cos(angle) * radius), TimeSpan.Zero));
                touchSequence.AddAction(touchContact.CreatePointerMove(CoordinateOrigin.Viewport, (int)(Math.Sin(angle) * radius), -(int)(Math.Cos(angle) * radius), TimeSpan.Zero));
                //touchSequence.AddAction(touchContact.CreatePointerMove(paintCanvas, 0, -radius, TimeSpan.Zero));
            }
            touchSequence.AddAction(touchContact.CreatePointerUp(PointerButton.TouchContact));


            session.PerformActions(new List<ActionSequence> { touchSequence });

            // Verify that the drawing operations took place
            WindowsElement undoButton = session.FindElementByAccessibilityId("UndoIcon");
            Assert.IsTrue(undoButton.Displayed);
            Assert.IsTrue(undoButton.Enabled);


        }

        private void Tap(int xOffset, int yOffset) {
            PointerInputDevice touch = new PointerInputDevice(PointerKind.Touch);
            ActionSequence touchSequence = new ActionSequence(touch, 0);
            touchSequence.AddAction(touch.CreatePointerMove(zoomInteractor, xOffset, yOffset, TimeSpan.Zero));
            touchSequence.AddAction(touch.CreatePointerDown(PointerButton.TouchContact));
            touchSequence.AddAction(touch.CreatePointerUp(PointerButton.TouchContact));
            session.PerformActions(new List<ActionSequence> { touchSequence });
        }
    }
}