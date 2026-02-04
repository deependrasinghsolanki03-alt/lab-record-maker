# 🚀 Lab Record Maker

**Lab Record Maker** is a powerful Visual Studio Code extension designed specifically for Computer Science students. It automates the tedious task of creating Lab Manuals and Practical Records by converting your C++ source code and its execution output into a **professionally formatted MS Word (.docx)** document.

---

## ✨ Features

- ✅ **Instant Word Generation:** Generate a complete lab record in seconds.
- ✅ **Professional Formatting:** Code is automatically wrapped in a clean **Gray Box** using the `Consolas` font for a professional look.
- ✅ **Smart Aim Detection:** Automatically detects the "Aim" or "Question" of the experiment from the first line of your code (comment).
- ✅ **Runtime Input Support:** Handles programs requiring user input (`cin`) by capturing the output based on your provided values.
- ✅ **Cover Page & Index Management:** Remembers your details to maintain a consistent cover page across all your experiments.

---

## ⚙️ Prerequisites

Before using this extension, ensure your system meets the following requirements:

1. **C++ Compiler (MinGW/G++):** - Verify by typing `g++ --version` in your terminal.
   - The compiler must be added to your system's **PATH** environment variable.
2. **Visual Studio Code:** The latest stable version.

---

## 📥 Installation

Since this extension is currently in development, you can install it manually using the `.vsix` file:

### **Method: Install from VSIX**
1. Download the **`lab-record-maker-0.0.1.vsix`** file from this repository.
2. Open Visual Studio Code.
3. Navigate to the **Extensions** view (`Ctrl + Shift + X`).
4. Click the **More Actions...** (three dots) at the top-right of the Extensions side bar.
5. Select **"Install from VSIX..."**.
6. Locate and select the downloaded `.vsix` file.

---

## 🚀 How to Use

### **1. Prepare Your Code**
Include the "Aim" of your experiment as a comment on the **very first line** of your `.cpp` file:
```cpp
// Write a program to find the largest of three numbers
#include <iostream>
using namespace std;
// ... your code


## 👨‍💻 Author

**Deependra Singh Solanki** *B.Tech Computer Science*