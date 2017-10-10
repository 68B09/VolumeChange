// VolumeChange.cpp : アプリケーションのエントリ ポイントを定義します。
//

#include "stdafx.h"
#include "VolumeChange.h"
#include <mmsystem.h>
#include <WinBase.h>
#include <Windows.h>
#include <Mmdeviceapi.h>
#include <Endpointvolume.h>

/*
[volume]
-1     ... ミュート
-2     ... ミュート解除
0～100 ... マスターボリューム 0～100%
*/
void SetMasterVolume(int volume)
{
	CoInitialize(NULL);

	IMMDeviceEnumerator *deviceEnumerator = NULL;
	HRESULT hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, CLSCTX_INPROC_SERVER, __uuidof(IMMDeviceEnumerator), (LPVOID *)&deviceEnumerator);
	IMMDevice *defaultDevice = NULL;

	hr = deviceEnumerator->GetDefaultAudioEndpoint(eRender, eConsole, &defaultDevice);
	deviceEnumerator->Release();
	deviceEnumerator = NULL;

	IAudioEndpointVolume *endpointVolume = NULL;
	hr = defaultDevice->Activate(__uuidof(IAudioEndpointVolume), CLSCTX_INPROC_SERVER, NULL, (LPVOID *)&endpointVolume);
	defaultDevice->Release();
	defaultDevice = NULL; 

	if( volume == -1 ){
		endpointVolume->SetMute(TRUE, NULL);
	}else if( volume == -2 ){
		endpointVolume->SetMute(FALSE, NULL);
	}else if( (volume >= 0 ) && (volume <= 100 ) ){
		endpointVolume->SetMasterVolumeLevelScalar(volume/100.0f, NULL);
	}

	endpointVolume->Release();

	CoUninitialize();
}

/*
VolumeChange option

[option]
muteon  ... ミュート
muteoff ... ミュート解除
0～100  ... マスターボリューム 0～100%
*/
int APIENTRY _tWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPTSTR lpCmdLine, int nCmdShow)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);

	int volume;

	if(_tcslen(lpCmdLine) > 0 ){
		if( _tcsicmp(lpCmdLine,_T("muteon")) == 0 ){
			volume = -1;
		}else if( _tcsicmp(lpCmdLine,_T("muteoff")) == 0 ){
			volume = -2;
		}else{
			volume = _tstoi(lpCmdLine);
		}

		SetMasterVolume(volume);
	}else{
		::MessageBox( NULL, _T("VolumeChange volume\nvolume=0～100 or muteon or muteoff"), _T("VolumeChange"), MB_OK);
	}

	return 0;
}
