/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT license.
 */

/* global Excel, Office */

const syntaxMap = {
  color: "color(int_row, char_column, string_color)",
  write: "write(int_row, char_column, string_value)",
  del: "del(int_row, char_column)",
  read: "read(int_row, char_column)",
};

const OPENAI_MODEL = "gpt-4o-mini";
const OPENAI_ENDPOINT = "https://api.openai.com/v1/chat/completions";
const OPENAI_API_KEY_PLACEHOLDER = "<YOUR_OPENAI_API_KEY>";
const MAX_REACT_LOOPS = 3;
const OPENAI_API_KEY = window.OPENAI_API_KEY || OPENAI_API_KEY_PLACEHOLDER;

let messagesElement;
let inputElement;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    initializeChat();
  }
});

function initializeChat() {
  const sideloadMsg = document.getElementById("sideload-msg");
  const chatApp = document.getElementById("chat-app");
  messagesElement = document.getElementById("messages");
  inputElement = document.getElementById("command-input");
  const form = document.getElementById("command-form");

  if (sideloadMsg) {
    sideloadMsg.style.display = "none";
  }
  if (chatApp) {
    chatApp.classList.remove("hidden");
  }

  form?.addEventListener("submit", async (event) => {
    event.preventDefault();
    if (!inputElement) {
      return;
    }
    const rawInput = inputElement.value;
    inputElement.value = "";
    await processCommand(rawInput);
  });

  inputElement?.addEventListener("keydown", (event) => {
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
      if (typeof form?.requestSubmit === "function") {
        form.requestSubmit();
      } else if (form) {
        form.dispatchEvent(new Event("submit", { bubbles: true, cancelable: true }));
      }
    }
  });

  appendMessage("system", "支持 color, write, del, read 四个命令；也可以输入自然语言让小助手判断并执行操作。");
}

async function processCommand(rawText) {
  const text = rawText.trim();
  if (!text) {
    return;
  }

  appendMessage("user", text);

  const parsed = parseCommand(text);
  if (!parsed) {
    await handleNaturalLanguage(text);
    return;
  }

  if (parsed.args === null) {
    appendMessage("assistant", syntaxReply(parsed.command));
    return;
  }

  const handler = commandHandlers[parsed.command];
  if (!handler) {
    await handleNaturalLanguage(text);
    return;
  }

  try {
    const response = await handler(parsed.args);
    appendMessage("assistant", response);
  } catch (error) {
    console.error(error);
    appendMessage("assistant", "操作失败，请稍后再试。");
  }
}

async function handleNaturalLanguage(text) {
  if (!OPENAI_API_KEY || OPENAI_API_KEY === OPENAI_API_KEY_PLACEHOLDER) {
    appendMessage("assistant", "未配置 OpenAI API Key，请在 window.OPENAI_API_KEY 中填入密钥。");
    return;
  }

  let observation = "尚未执行任何操作。";

  for (let loop = 0; loop < MAX_REACT_LOOPS; loop += 1) {
    const reactResult = await askOpenAiForPlan({
      userInput: text,
      observation,
      loopIndex: loop + 1,
      remainingLoops: MAX_REACT_LOOPS - loop,
    });

    if (reactResult.error) {
      appendMessage("assistant", reactResult.error);
      return;
    }

    appendMessage("assistant", `思考：${reactResult.reasoning || "无"}`);

    if (!reactResult.requiresExcelAction) {
      appendMessage("assistant", reactResult.reply || "好的，我理解了。");
      return;
    }

    const command = reactResult.command;
    if (!command) {
      appendMessage("assistant", "未能识别需要执行的 Excel 命令。");
      return;
    }

    const handler = commandHandlers[command];
    if (!handler) {
      appendMessage("assistant", `目前不支持 ${command} 命令。`);
      return;
    }

    try {
      const args = (reactResult.args ?? []).map((arg) => `${arg}`.trim());
      const actionResponse = await handler(args);

      const isSyntaxHint = actionResponse === syntaxMap[command];
      const messageText = reactResult.reply || actionResponse;
      appendMessage("assistant", messageText);

      observation = actionResponse;

      if (isSyntaxHint) {
        return;
      }

      if (reactResult.done) {
        return;
      }
    } catch (error) {
      console.error(error);
      appendMessage("assistant", "执行 Excel 操作时出错，请稍后再试。");
      return;
    }
  }

  appendMessage("assistant", "已达到最大推理次数，暂时停止。");
}

async function askOpenAiForPlan({ userInput, observation, loopIndex, remainingLoops }) {
  const systemPrompt = `你是 Excel 扩展的智能助手，遵循 Excel JavaScript API，关联文档 https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=excel-js-preview。`;
  const instructionPrompt = `
用户输入: ${userInput}
当前观察: ${observation}
本轮编号: ${loopIndex}
剩余迭代次数: ${remainingLoops}

请仅输出符合 JSON 结构的内容，字段如下：
- reasoning: 思考过程，描述如何判断用户需求并链接 Excel 操作或回复；
- requiresExcelAction: 布尔，是否需要执行 Excel 命令；
- command: color/write/del/read 之一（若 requiresExcelAction 为 false 可为空）；
- args: 字符串数组，依次对应命令参数；
- reply: 用于展示给用户的自然语言回复；
- done: 布尔，指示是否已完成用户目标，若 true 可终止循环。

必须确保 JSON 可解析，配置好命令所需参数，若不需要 Excel 操作直接提供 reply。`;

  try {
    const response = await fetch(OPENAI_ENDPOINT, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${OPENAI_API_KEY}`,
      },
      body: JSON.stringify({
        model: OPENAI_MODEL,
        messages: [
          { role: "system", content: systemPrompt },
          { role: "user", content: instructionPrompt },
        ],
        temperature: 0.2,
      }),
    });

    if (!response.ok) {
      const payload = await response.json().catch(() => ({}));
      const message = payload.error?.message || "调用 OpenAI 接口失败。";
      return { error: message };
    }

    const data = await response.json();
    const content = data.choices?.[0]?.message?.content ?? "";
    const parsed = safeParseJson(content);
    const normalized = normalizeAiPayload(parsed);
    return normalized;
  } catch (error) {
    console.error(error);
    return { error: "与 OpenAI 通信失败，请检查网络或密钥配置。" };
  }
}

function normalizeAiPayload(payload) {
  const requiresExcelAction = Boolean(payload.requiresExcelAction);
  const commandRaw = payload.command;
  const command = typeof commandRaw === "string" ? commandRaw.toLowerCase() : null;

  const args = Array.isArray(payload.args)
    ? payload.args.map((value) => `${value}`)
    : [];

  return {
    reasoning: `${payload.reasoning ?? ""}`.trim(),
    requiresExcelAction,
    command,
    args,
    reply: `${payload.reply ?? ""}`.trim(),
    done: payload.done === true,
  };
}

function safeParseJson(text) {
  const start = text.indexOf("{");
  const end = text.lastIndexOf("}");
  if (start === -1 || end === -1) {
    throw new Error("OpenAI 未返回有效 JSON。");
  }
  const snippet = text.slice(start, end + 1);
  return JSON.parse(snippet);
}

function parseCommand(input) {
  const match = input.match(/^([a-z]+)\s*\(\s*([\s\S]*?)\s*\)$/i);
  if (!match) {
    return null;
  }
  const command = match[1].toLowerCase();
  const argsString = match[2];
  const args = argsString === "" ? [] : splitArguments(argsString);
  if (args === null) {
    return { command, args: null };
  }
  return { command, args };
}

function splitArguments(text) {
  const args = [];
  let current = "";
  let inSingle = false;
  let inDouble = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];

    if (char === "'" && !inDouble) {
      inSingle = !inSingle;
      continue;
    }

    if (char === '"' && !inSingle) {
      inDouble = !inDouble;
      continue;
    }

    if (char === "," && !inSingle && !inDouble) {
      args.push(current.trim());
      current = "";
      continue;
    }

    current += char;
  }

  if (inSingle || inDouble) {
    return null;
  }

  if (current || text.endsWith(",")) {
    args.push(current.trim());
  }

  return args.map((arg) => stripQuotes(arg));
}

function stripQuotes(value) {
  const trimmed = value.trim();
  if (trimmed.length >= 2) {
    const first = trimmed[0];
    const last = trimmed[trimmed.length - 1];
    if ((first === "'" && last === "'") || (first === '"' && last === '"')) {
      return trimmed.slice(1, -1);
    }
  }
  return trimmed;
}

function syntaxReply(command) {
  return syntaxMap[command] ?? "暂不支持该语法";
}

const commandHandlers = {
  color: async (args) => {
    if (args.length !== 3) {
      return syntaxReply("color");
    }

    const row = parseRow(args[0]);
    const column = normalizeColumn(args[1]);
    const colorValue = args[2].trim();

    if (!row || !column || !colorValue) {
      return syntaxReply("color");
    }

    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(`${column}${row}`);
      range.format.fill.color = colorValue;
      await context.sync();
    });

    return `已将 ${column}${row} 设置为 ${colorValue} 颜色。`;
  },
  write: async (args) => {
    if (args.length !== 3) {
      return syntaxReply("write");
    }

    const row = parseRow(args[0]);
    const column = normalizeColumn(args[1]);
    const value = args[2];

    if (!row || !column || value === undefined) {
      return syntaxReply("write");
    }

    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(`${column}${row}`);
      range.values = [[value]];
      await context.sync();
    });

    return `已写入 ${column}${row}：${value}`;
  },
  del: async (args) => {
    if (args.length !== 2) {
      return syntaxReply("del");
    }

    const row = parseRow(args[0]);
    const column = normalizeColumn(args[1]);

    if (!row || !column) {
      return syntaxReply("del");
    }

    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(`${column}${row}`);
      range.clear(Excel.ClearApplyTo.contents);
      await context.sync();
    });

    return `已清除 ${column}${row} 的内容。`;
  },
  read: async (args) => {
    if (args.length !== 2) {
      return syntaxReply("read");
    }

    const row = parseRow(args[0]);
    const column = normalizeColumn(args[1]);

    if (!row || !column) {
      return syntaxReply("read");
    }

    let displayValue = "无值";

    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(`${column}${row}`);
      range.load("values");
      await context.sync();
      const cellValue = range.values?.[0]?.[0];
      if (cellValue !== undefined && cellValue !== null && cellValue !== "") {
        displayValue = `${cellValue}`;
      }
    });

    return `读取 ${column}${row}：${displayValue}`;
  },
};

function parseRow(value) {
  const number = Number(value);
  if (!Number.isInteger(number) || number <= 0) {
    return null;
  }
  return number;
}

function normalizeColumn(value) {
  const column = value?.trim().toUpperCase();
  if (!column || !/^[A-Z]+$/.test(column)) {
    return null;
  }
  return column;
}

function appendMessage(role, text) {
  if (!messagesElement) {
    return;
  }

  const messageItem = document.createElement("div");
  messageItem.className = `message ${role}`;
  messageItem.textContent = text;

  messagesElement.appendChild(messageItem);
  messagesElement.scrollTop = messagesElement.scrollHeight;
}
