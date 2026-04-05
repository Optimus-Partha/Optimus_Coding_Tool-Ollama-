# Optimus Coding Tool
This is a coding helper tool similar to Claude Code but running locally using Ollama models.

This code is using `ollama/qwen3:4b`. To change the model check config.py - `model`:.

To run the tool, use `python coding_tool.py`

## Configuration
- **Model**: Use `/model [model]` to set the model (e.g., `ollama/qwen3:4b`, `gpt-4o`)
- **Output Directory**: All generated files (code, scripts, Excel files, etc.) MUST be saved in `./test/` directory
- **Output Rules**: Generated files include:
  - Code files
  - Helper scripts
  - Excel files
  - Temporary files
  - Any other artifacts

## Usage
```bash
python coding_tool.py
# With specific model
python coding_tool.py --model ollama/qwen3:4b
# With output directory
python coding_tool.py --output ./test
```

## Multi-Provider Support
Supports Anthropic, OpenAI, Gemini, Kimi, Qwen, Zhipu, DeepSeek, Ollama, LM Studio models

## Excel Automation
Use `/excel` command to automate Excel operations

## Security
- All file operations include permission gating
- Bash commands are filtered for safety
- No sensitive information stored in code
