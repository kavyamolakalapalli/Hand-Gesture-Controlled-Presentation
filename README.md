# Hand-Gesture-Controlled-Presentation

In this, I developed a Hand Gesture-Controlled Presentation System that combines computer vision with user-friendly interaction. It leverages OpenCV, hand gesture recognition, and multimedia tools to allow users to control PowerPoint presentations through gestures, without needing a physical device like a mouse or remote.

Key Components and Features of the Project

 1.PowerPoint-to-Image Conversion:
                The presentation begins by converting PowerPoint slides into PNG images. This is achieved using win32com.client, which opens the PowerPoint file, exports each slide as a PNG, and stores them in an output folder. This image format makes it easier for OpenCV to process and display each slide.

 2.Hand Gesture Detection:
                Using OpenCV and cvzone's HandDetector module, the system detects the user's hand gestures via a webcam feed. The detector identifies specific hand gestures and maps them to presentation commands, like moving to the next slide, going back, drawing annotations, and undoing actions.

 3.Each gesture corresponds to specific functions:
                Swipe Left/Right: Move to the previous or next slide.
                Pointing: Acts as a pointer, highlighting specific areas on the slide.
                Drawing: Allows users to annotate on the slide with a virtual pen.
                Undo: Removes the last annotation.
                Zoom In/Out: Enables zooming, useful for emphasizing slide sections.

4.Real-Time Slide Presentation and Annotation:
                The processed images are displayed in full screen, and annotations are rendered on top of each slide in real time. Users can add or remove annotations as needed during the presentation.
                The system visually marks the threshold line for gestures, making it intuitive to know where to place the hand for gesture recognition.

5.Video Recording and Audio Extraction:
                The application also includes a recording feature to capture the presentation. Users can toggle recording with a keypress, saving the session as a video file.Once recording stops, the system extracts audio from the video file using moviepy.editor, enabling users to review both video and audio aspects of their presentation later.

6.Automated Slide Cleanup:
                After the presentation ends, all slides are automatically deleted from the folder to keep the workspace organized.

Technologies and Libraries Used:

        OpenCV: For image processing, camera handling, and gesture visualization.
        CVZone: For easy hand tracking and gesture recognition through the HandTrackingModule.
        MoviePy: For video editing and audio extraction, enabling users to save their presentation as a video file.
        Win32com: To automate PowerPoint operations, converting slides to PNGs.