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
// ȫ�ֱ���
BOOL g_commandExecuted = FALSE;  // ��������Ƿ���ִ��
BOOL g_executionSuccess = FALSE; // ��������Ƿ�ִ�гɹ�
BOOL g_stopEnumeration = FALSE;  // ����Ƿ�ֹͣö��
// ��־�������򻯰棩
void log_message(const wchar_t* message) {
	wprintf(L"[LOG] %s\n", message);
}
// ���������ַ���
//wchar_t* build_command(const wchar_t* path) {
//	size_t len = wcslen(path) + 20;
//	wchar_t* cmd = (wchar_t*)malloc(len * sizeof(wchar_t));
//	if (cmd) {
//		swprintf_s(cmd, len, L"cmd /c start  \"\" \"%s\"", path);
//	}
//	return cmd;
//}
wchar_t* build_command(const wchar_t* path) {
	// ֱ�Ӹ���·��������ʹ��cmd /c start
	size_t len = wcslen(path) + 1;
	wchar_t* cmd = (wchar_t*)malloc(len * sizeof(wchar_t));
	if (cmd) {
		wcscpy_s(cmd, len, path);
	}
	return cmd;
}
// RPC�󶨽���
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
// ͨ��WMIִ������
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
		// ����WMI��λ��
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
		// ���ӵ�WMI�����ռ�
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
		// ���ô���ȫ����
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
		// ��ȡWin32_Process��
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
		// ��ȡCreate����
		hr = process->GetMethod(_bstr_t(L"Create"), 0, &method, NULL);
		if (FAILED(hr)) {
			log_message(L"Failed to get Create method");
			break;
		}
		// �����������
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
		// ִ�з���
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
	// ������Դ
	if (in_params) in_params->Release();
	if (method) method->Release();
	if (process) process->Release();
	if (services) services->Release();
	if (locator) locator->Release();
	if (out_params) out_params->Release();
	CoUninitialize();
	return success;
}
// �෽������ִ��
BOOL execute_command(const wchar_t* command) {
	log_message(L"Starting multi-method command execution");

	// ����1: WMIִ��
	log_message(L"Attempting method 1: WMI execution");
	if (execute_via_wmi(command)) {
		log_message(L"Method 1 successful: WMI execution");
		return TRUE;
	}

	// ����2: ֱ��APIִ��
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
	// ����Ѿ�ֹͣö�٣�ֱ�ӷ���FALSE
	if (g_stopEnumeration) {
		return FALSE;
	}

	// ��������Ѿ�ִ�гɹ���ֹͣö��
	if (g_executionSuccess) {
		g_stopEnumeration = TRUE;
		return FALSE;  // ��ȷ��ֹö��
	}

	// ���������δִ�У�����ִ��
	if (!g_commandExecuted) {
		log_message(L"Executing command via window enumeration");
		const wchar_t* target_path = L"C:\\Windows\\System32\\calc.exe";
		wchar_t* command = build_command(target_path);
		if (command) {
			g_commandExecuted = TRUE; // ��������ѳ���ִ��
			if (execute_command(command)) {
				g_executionSuccess = TRUE; // ���ִ�гɹ�
				g_stopEnumeration = TRUE;  // ���ֹͣö��
				free(command);
				return FALSE;  // �ɹ���������ֹö��
			}
			free(command);
		}
	}

	// ����ö����������
	return TRUE;
}
int main() {
	SetConsoleOutputCP(CP_UTF8);


	// ö�ٴ���
	if (!EnumWindows(EnumWindowsProc, 0)) {
		DWORD error = GetLastError();
		if (error == 0) {
			log_message(L"EnumWindows stopped by callback (normal termination)");
		}
		else {
			wprintf(L"EnumWindows failed. Error: %d\n", error);
		}
	}

	// ���ִ�н��
	if (g_executionSuccess) {
		log_message(L"Execution successful via window enumeration");
	}
	else {
		log_message(L"Execution failed");
	}

	// ����COM
	CoUninitialize();
	return g_executionSuccess ? 0 : 1;
}