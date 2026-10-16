import { zhCNCalendar } from './zh-CN/calendar';
import { zhCNChat } from './zh-CN/chat';
import { zhCNCore } from './zh-CN/core';
import { zhCNFiles } from './zh-CN/files';
import { zhCNSettings } from './zh-CN/settings';
import { zhCNShell } from './zh-CN/shell';

/** 简体中文。 */
export const zhCN = {
	...zhCNShell,
	...zhCNChat,
	...zhCNCalendar,
	...zhCNFiles,
	...zhCNSettings,
	...zhCNCore,
};
