#include <windows.h>
#include <stdio.h>
#include <oleauto.h>
#include <wbemidl.h>
#include <rpc.h>
#include <comdef.h>
#pragma comment(lib, "rpcrt4.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "wbemuuid.lib")
// 全局变量
BOOL g_commandExecuted = FALSE;  // 标记命令是否已执行
BOOL g_executionSuccess = FALSE; // 标记命令是否执行成功
BOOL g_stopEnumeration = FALSE;  // 标记是否停止枚举
// 日志函数（简化版）
void log_message(const wchar_t* message) {
	wprintf(L"[LOG] %s\n", message);
}
// 构建命令字符串
//wchar_t* build_command(const wchar_t* path) {
//	size_t len = wcslen(path) + 20;
//	wchar_t* cmd = (wchar_t*)malloc(len * sizeof(wchar_t));
//	if (cmd) {
//		swprintf_s(cmd, len, L"cmd /c start  \"\" \"%s\"", path);
//	}
//	return cmd;
//}
wchar_t* build_command(const wchar_t* path) {
	// 直接复制路径，不再使用cmd /c start
	size_t len = wcslen(path) + 1;
	wchar_t* cmd = (wchar_t*)malloc(len * sizeof(wchar_t));
	if (cmd) {
		wcscpy_s(cmd, len, path);
	}
	return cmd;
}
// RPC绑定建立
BOOL establish_rpc_binding(RPC_BINDING_HANDLE* binding, const wchar_t* service) {
	RPC_WSTR binding_string = NULL;
	RPC_WSTR endpoint = (RPC_WSTR)malloc(256 * sizeof(wchar_t));
	if (!endpoint) return FALSE;
	swprintf_s((wchar_t*)endpoint, 256, L"\\pipe\\%s", service);
	RPC_STATUS status = RpcStringBindingComposeW(
		NULL,
		(RPC_WSTR)L"ncalrpc",
		NULL,
		endpoint,
		NULL,
		&binding_string
	);
	if (status != RPC_S_OK) {
		log_message(L"RPC binding composition failed");
		free(endpoint);
		return FALSE;
	}
	status = RpcBindingFromStringBindingW(binding_string, binding);
	RpcStringFreeW(&binding_string);
	free(endpoint);
	if (status == RPC_S_OK) {
		log_message(L"RPC binding established successfully");
		return TRUE;
	}
	log_message(L"RPC binding from string failed");
	return FALSE;
}
// 通过WMI执行命令
BOOL execute_via_wmi(const wchar_t* command) {
	HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
	if (FAILED(hr)) {
		log_message(L"COM initialization failed");
		return FALSE;
	}
	IWbemLocator* locator = NULL;
	IWbemServices* services = NULL;
	IWbemClassObject* process = NULL;
	IWbemClassObject* out_params = NULL;
	IWbemClassObject* method = NULL;
	IWbemClassObject* in_params = NULL;
	BOOL success = FALSE;
	do {
		// 创建WMI定位器
		hr = CoCreateInstance(
			CLSID_WbemLocator,
			0,
			CLSCTX_INPROC_SERVER,
			IID_IWbemLocator,
			(LPVOID*)&locator
		);
		if (FAILED(hr)) {
			log_message(L"Failed to create WbemLocator");
			break;
		}
		// 连接到WMI命名空间
		hr = locator->ConnectServer(
			_bstr_t(L"ROOT\\CIMV2"),
			NULL,
			NULL,
			0,
			NULL,
			0,
			0,
			&services
		);
		if (FAILED(hr)) {
			log_message(L"Failed to connect to WMI namespace");
			break;
		}
		// 设置代理安全级别
		hr = CoSetProxyBlanket(
			services,
			RPC_C_AUTHN_WINNT,
			RPC_C_AUTHZ_NONE,
			NULL,
			RPC_C_AUTHN_LEVEL_CALL,
			RPC_C_IMP_LEVEL_IMPERSONATE,
			NULL,
			EOAC_NONE
		);
		if (FAILED(hr)) {
			log_message(L"Failed to set proxy blanket");
			break;
		}
		// 获取Win32_Process类
		hr = services->GetObject(
			_bstr_t(L"Win32_Process"),
			0,
			NULL,
			&process,
			NULL
		);
		if (FAILED(hr)) {
			log_message(L"Failed to get Win32_Process class");
			break;
		}
		// 获取Create方法
		hr = process->GetMethod(_bstr_t(L"Create"), 0, &method, NULL);
		if (FAILED(hr)) {
			log_message(L"Failed to get Create method");
			break;
		}
		// 设置输入参数
		hr = method->SpawnInstance(0, &in_params);
		if (FAILED(hr)) {
			log_message(L"Failed to spawn method instance");
			break;
		}
		VARIANT var_command;
		VariantInit(&var_command);
		var_command.vt = VT_BSTR;
		var_command.bstrVal = SysAllocString(command);
		hr = in_params->Put(_bstr_t(L"CommandLine"), 0, &var_command, 0);
		VariantClear(&var_command);
		if (FAILED(hr)) {
			log_message(L"Failed to put command line parameter");
			break;
		}
		// 执行方法
		hr = services->ExecMethod(
			_bstr_t(L"Win32_Process"),
			_bstr_t(L"Create"),
			0,
			NULL,
			in_params,
			&out_params,
			NULL
		);
		if (SUCCEEDED(hr)) {
			log_message(L"WMI process creation successful");
			success = TRUE;
		}
		else {
			log_message(L"WMI process creation failed");
		}
	} while (0);
	// 清理资源
	if (in_params) in_params->Release();
	if (method) method->Release();
	if (process) process->Release();
	if (services) services->Release();
	if (locator) locator->Release();
	if (out_params) out_params->Release();
	CoUninitialize();
	return success;
}
// 多方法命令执行
BOOL execute_command(const wchar_t* command) {
	log_message(L"Starting multi-method command execution");

	// 方法1: WMI执行
	log_message(L"Attempting method 1: WMI execution");
	if (execute_via_wmi(command)) {
		log_message(L"Method 1 successful: WMI execution");
		return TRUE;
	}

	// 方法2: 直接API执行
	log_message(L"Attempting method 2: Direct API execution");
	STARTUPINFOW si = { sizeof(si) };
	PROCESS_INFORMATION pi = { 0 };
	si.dwFlags = STARTF_USESHOWWINDOW;
	si.wShowWindow = SW_SHOW;

	if (CreateProcessW(
		NULL,
		(LPWSTR)command,
		NULL,
		NULL,
		FALSE,
		CREATE_NO_WINDOW,
		NULL,
		NULL,
		&si,
		&pi)) {
		log_message(L"Method 2 successful: Direct API execution");
		CloseHandle(pi.hProcess);
		CloseHandle(pi.hThread);
		return TRUE;
	}

	log_message(L"All execution methods failed");
	return FALSE;
}
// Callback function for EnumWindows
BOOL CALLBACK EnumWindowsProc(HWND hwnd, LPARAM lParam) {
	// 如果已经停止枚举，直接返回FALSE
	if (g_stopEnumeration) {
		return FALSE;
	}

	// 如果命令已经执行成功，停止枚举
	if (g_executionSuccess) {
		g_stopEnumeration = TRUE;
		return FALSE;  // 正确终止枚举
	}

	// 如果命令尚未执行，尝试执行
	if (!g_commandExecuted) {
		log_message(L"Executing command via window enumeration");
		const wchar_t* target_path = L"C:\\Windows\\System32\\calc.exe";
		wchar_t* command = build_command(target_path);
		if (command) {
			g_commandExecuted = TRUE; // 标记命令已尝试执行
			if (execute_command(command)) {
				g_executionSuccess = TRUE; // 标记执行成功
				g_stopEnumeration = TRUE;  // 标记停止枚举
				free(command);
				return FALSE;  // 成功后立即终止枚举
			}
			free(command);
		}
	}

	// 继续枚举其他窗口
	return TRUE;
}
int main() {
	SetConsoleOutputCP(CP_UTF8);


	// 枚举窗口
	if (!EnumWindows(EnumWindowsProc, 0)) {
		DWORD error = GetLastError();
		if (error == 0) {
			log_message(L"EnumWindows stopped by callback (normal termination)");
		}
		else {
			wprintf(L"EnumWindows failed. Error: %d\n", error);
		}
	}

	// 检查执行结果
	if (g_executionSuccess) {
		log_message(L"Execution successful via window enumeration");
	}
	else {
		log_message(L"Execution failed");
	}

	// 清理COM
	CoUninitialize();
	return g_executionSuccess ? 0 : 1;
}