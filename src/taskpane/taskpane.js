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

  appendMessage("system", "支持的命令包括 color、write、del、read；如语法不对会返回格式说明。");
}

async function processCommand(rawText) {
  const text = rawText.trim();
  if (!text) {
    return;
  }

  appendMessage("user", text);

  const parsed = parseCommand(text);
  if (!parsed) {
    appendMessage("assistant", "暂不支持该语法");
    return;
  }

  const handler = commandHandlers[parsed.command];

  if (!handler) {
    appendMessage("assistant", "暂不支持该语法");
    return;
  }

  if (parsed.args === null) {
    appendMessage("assistant", syntaxReply(parsed.command));
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

  for (let i = 0; i < text.length; i++) {
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
