# 🖐️ Hand Gesture Controlled PowerPoint Presentation

<div align="center"> 
    <img src="p1.png" alt="Hand Gesture PPT Control" width="600"/>
</div> 

Control your PowerPoint presentations using simple hand gestures! This project offers an **interactive, hands-free** experience by leveraging computer vision and gesture detection.

---

## 🚀 Features

- ✋ Real-time hand detection using OpenCV and MediaPipe  
- 👆 Intuitive gesture recognition via cvzone  
- 🎞️ Seamless slide control (Next/Previous) in MS PowerPoint  
- 🧠 Powered by machine learning-based hand landmark tracking  

---

## 🛠️ Tech Stack

- **Language**: Python  
- **Libraries**: OpenCV, cvzone, MediaPipe, pywin32  
- **Platform**: Windows (PowerPoint automation requires MS Office)  

---

## 📦 Installation & Setup

Follow these steps to set up the project on your system:

1. **Clone the repository**:
   ```bash
   git clone https://github.com/your-username/hand-gesture-ppt-control.git
   cd hand-gesture-ppt-control
   ```

2. **Create and activate a virtual environment (recommended)**:
   ```bash
   python -m venv venv
   .\venv\Scripts\activate   # For Windows
   ```

3. **Install all required libraries**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Edit the PowerPoint path in `Code.py`**:
   Open the file and update this line:
   ```python
   Presentation = Application.Presentations.Open("C:\\Path\\To\\Your\\Presentation.pptx")
   ```
   Replace it with the actual path of your PowerPoint file on your machine.

---

## ▶️ How to Run

1. Make sure your webcam is connected and MS PowerPoint is installed.
2. Run the script:
   ```bash
   python Code.py
   ```
3. Use these gestures:
   - ✋ All five fingers up → **Next slide**
   - ☝ Only thumb up → **Previous slide**
4. Press `Q` to quit the application.

---

## 📸 Demo

> 
<div align="center"> 
    <img src="p1.png" alt="Hand Gesture PPT Control" width="600"/>
</div> 

---

## ⚠️ Notes

- Works best with a **single hand** and under **good lighting conditions**.
- Ensure you are facing the webcam directly for best gesture detection.
- This was tested on **Windows 10 with MS PowerPoint 2016**.

---

## 🙏 Acknowledgments

- [cvzone](https://github.com/cvzone/cvzone) – Simplified computer vision utilities  
- [MediaPipe](https://mediapipe.dev/) – Hand tracking framework  
- Microsoft Office – PowerPoint automation via `pywin32`  

---

## 📬 Contact

Made with ❤️ by **Devashish Soni**  
🔗 [LinkedIn](https://www.linkedin.com/in/devashish-soni/)  
For issues, suggestions, or collaboration, feel free to raise an issue in the repository.
