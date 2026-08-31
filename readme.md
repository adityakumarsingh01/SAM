<div align="center">
  <img src="SAM.png" alt="SAM Logo" width="200" />
  <h1>SAM (Smart Adaptive Mentor)</h1>
  <p><strong>The Ultimate Autonomous Cross-Platform AI Assistant</strong></p>
  <p><em>Engineered by Aditya Kumar Singh</em></p>
  
  [![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
  [![Python](https://img.shields.io/badge/Python-3.11+-yellow.svg)](https://www.python.org/)
  [![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)]()
</div>

<br />

## 📖 Executive Summary
**SAM (Smart Adaptive Mentor)** is an advanced, highly autonomous personal AI assistant built on the Gemini Live API. Designed to function as an extension of the user's digital life, SAM possesses the ability to hear, see, understand, and directly interact with computer systems across Windows, macOS, and Linux. 

Unlike traditional voice assistants, SAM features **Affective Dialog** (detecting and matching human emotion), **Contextual Proactivity** (knowing when to speak and when to stay silent), and a **Zero-Friction Plugin Architecture** that allows developers to expand its capabilities without altering core engine files. SAM is not just a chatbot; it is a "God-level" copilot capable of high-level task planning, real-time screen vision, and autonomous desktop control.

This project serves as a comprehensive demonstration of modern AI integration, system-level automation, and scalable software architecture.

---

## 🏗️ System Architecture & Technologies

SAM is built with a focus on modularity, low latency, and infinite extensibility. 

- **Core Intelligence**: Google Gemini Live API (for real-time, low-latency multi-modal processing).
- **Frontend / HUD**: PyQt6 (providing a dynamic, hardware-accelerated Heads-Up Display with real-time waveform rendering).
- **System Integration**: `pyautogui`, `win32com` (for deep OS-level and application-level automation, including direct Microsoft Excel COM manipulation).
- **State Management**: Sliding-window context compression, allowing for infinite session lengths without token overflow.
- **Plugin Engine**: Dynamic Python module loading that auto-discovers and registers new capabilities at runtime securely.

---

## 🚀 Key Capabilities

SAM possesses a vast array of built-in capabilities, making it one of the most powerful local assistants available:

### 🧠 Advanced AI & Memory
* **Affective Dialog**: SAM listens to the tone and emotion in your voice and dynamically adjusts its responses (e.g., matching excitement or responding calmly to fatigue).
* **Proactive Silence**: Advanced audio filtering prevents SAM from responding to background chatter, TVs, or phone calls. No wake-word required.
* **Persistent Memory**: SAM deeply remembers projects, preferences, and personal context across sessions, summarizing its memory daily to optimize context limits.

### 💻 System & Desktop Automation
* **Excel Copilot**: Interacts directly with Microsoft Excel's memory via COM. SAM can format cells, read massive tables, and perform lightning-fast 2D array batch updates without hallucinatory errors.
* **Full Desktop Control**: Automates mouse, keyboard, window management, and taskbar operations.
* **Hardware Telemetry**: Real-time monitoring of CPU, RAM, GPU, and thermal statistics with localized voice alerts.
* **App Management**: Intelligently launches applications or brings existing instances to the foreground to prevent cloning.

### 🌐 Web, Media, & Productivity
* **Spotify Automation**: Native integration to search, play, pause, and control music playback via URI handlers.
* **Multi-Mode Web Search**: Capable of executing live news aggregation, price tracking, and deep research using Gemini Grounded search with DuckDuckGo fallbacks.
* **Code & Developer Agent**: Inline code review, script generation, and autonomous debugging capabilities.
* **Proactive Reminders**: Integrates directly with OS-native scheduling (Windows Task Scheduler / systemd) to push context-aware reminders.

---

## 🧩 The Plugin Architecture

SAM's architecture is finalized. All new features are added via the **Plugin System** without ever touching the core engine. 

### How it works:
1. Write a single `.py` file detailing your Gemini tool schema and logic (using `plugins/_template.py`).
2. Drop the file into the `plugins/` directory.
3. Restart SAM.

The engine automatically discovers the plugin, validates it, registers it with the Gemini Live session, and populates it in the UI's Plugin Manager. **Safety is guaranteed**: a broken plugin will never crash SAM; it will simply be isolated and marked as "BROKEN" in the dashboard.

---

## ⚡ Quick Start & Installation

To run SAM locally on your machine:

```bash
# 1. Clone the repository
git clone https://github.com/adityakumarsingh01/SAM.git
cd SAM

# 2. Install dependencies
pip install -r requirements.txt

# 3. Launch the Assistant
python main.py
```
> **Note**: Ensure you have a valid Gemini API key. Upon first boot, SAM's setup wizard will guide you through entering your credentials and personalizing your assistant.

---

## 🗂️ Repository Structure

```text
SAM/
├── main.py                   # Core event loop, tool dispatch, and audio I/O
├── ui.py                     # PyQt6 hardware-accelerated HUD and Plugin Manager
├── plugins/                  # Drop-in directory for infinite extensibility
├── actions/                  # Core capability scripts (System, Web, Files, Desktop)
├── memory/                   # Persistent storage and memory compression engine
├── core/                     # Plugin loader, core routing rules, and LLM prompts
└── config/                   # API keys and localized user settings
```

---

## 🔮 Future Scope & Roadmap

- **Local LLM Fallback**: Integration with local models (e.g., Llama 3 via Ollama) for offline functionality and privacy-critical tasks.
- **Multi-Agent Swarms**: Allowing SAM to spawn background sub-agents for heavy research tasks while maintaining the main conversation.
- **Smart Home Integration**: Expanding network capabilities to natively control IoT devices via HomeAssistant APIs.

---

## 📜 License

This project is open-source and released under the **MIT License**. You are free to use, modify, distribute, and commercialize this software.

---

## 👨‍💻 Connect with the Creator

Engineered with passion by **Aditya Kumar Singh**. I am constantly exploring the bleeding edge of AI, automation, and software architecture.

- 🌐 **Portfolio**: [Visit my Website](https://portfolio-ecru-one-nzr8n36bhi.vercel.app/)
- 💼 **LinkedIn**: [Connect with me](https://www.linkedin.com/in/aditya-kumar-singh-990377291/)
- ✉️ **Email**: [adityasingh81201@gmail.com](mailto:adityasingh81201@gmail.com)
