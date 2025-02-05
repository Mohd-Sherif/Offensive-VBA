# Offensive VBA

A collection of offensive VBA (Visual Basic for Applications) scripts designed for red teaming, penetration testing, and security research. This repository also includes a **VBA Learning Guide** for those new to VBA programming.

---

## Introduction

This repository focuses on **offensive VBA scripts** that can be used in red teaming and penetration testing scenarios. These scripts demonstrate how VBA can be leveraged for payload delivery, evasion, and post-exploitation in authorized environments.

Additionally, a **VBA Learning Guide** is included for those who want to learn or brush up on VBA programming. This guide provides foundational knowledge and practical examples that are essential for creating offensive VBA scripts.

---

## Features

### VBA Learning Guide
- A comprehensive collection of VBA examples for beginners.
- Covers loops, conditions, functions, error handling, and more.
- Interactive menu for easy navigation.
- **Program Execution Demo**: Learn how to execute external programs, run PowerShell commands, and capture output using VBA.

### Offensive VBA Scripts (Coming Soon!)
- **Payload Delivery**: Examples of delivering payloads via VBA macros.
- **Evasion Techniques**: Scripts to bypass security controls.
- **Post-Exploitation**: Tools for maintaining access and executing commands.
- **Real-World Scenarios**: Practical examples for red teaming.

---

## Getting Started

### Prerequisites
- Microsoft Excel or any Office application that supports VBA.
- **Authorization**: Only use these scripts in environments where you have explicit permission.

### How to Use
1. Clone this repository or download the VBA files.
2. Open the VBA file in Excel (or another Office application).
3. Press `Alt + F11` to open the VBA editor.
4. Run the desired script or explore the **VBA Learning Guide** using the `MainMenu` procedure.

---

## VBA Learning Guide

The **VBA Learning Guide** is designed to help you master the basics of VBA programming, which are essential for creating offensive scripts. It includes interactive examples and a **Program Execution Demo** to teach you how to execute external programs and commands.

### Program Execution Demo
This script demonstrates how to execute external programs, run PowerShell commands, and capture output using VBA. It includes a Main Menu for selecting examples and detailed error handling.

#### Features
- Run external programs (e.g., Notepad, Calculator).
- Execute PowerShell commands from VBA.
- Capture and display command output (e.g., `ipconfig`).
- User-friendly menu for selecting examples.

#### Usage
1. Import the `VBA_Program_Execution_Demo.vba` file into your VBA project.
2. Run the `AutoOpen` subroutine to start the Main Menu.
3. Choose an example to execute and view the results.

#### Examples
- **Run Notepad**: Opens Notepad using the `Shell` function.
- **Run Notepad via PowerShell**: Executes Notepad in memory using PowerShell.
- **Save Process List**: Saves a list of running processes to a file using PowerShell.
- **Open Calculator**: Opens Calculator asynchronously.
- **Run ipconfig**: Executes `ipconfig` and displays the output.

---

## Offensive VBA Scripts (Coming Soon!)

This section will include advanced scripts for:
- Macro-based payload delivery.
- Bypassing security controls (e.g., AMSI, antivirus).
- Executing commands and maintaining access.

Stay tuned for updates!

---

## Disclaimer

This repository is for **educational and ethical purposes only**. The offensive VBA scripts are intended for use in authorized penetration testing and red teaming activities. Do not use these scripts for malicious purposes. The author is not responsible for any misuse or damage caused by the scripts in this repository.

**Always obtain proper authorization before testing.**

---
