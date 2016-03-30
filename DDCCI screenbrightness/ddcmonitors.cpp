//#include "stdafx.h"
#include "ddcmonitors.h"
//#include <windows.h>
#include <lowlevelmonitorconfigurationapi.h>
#include <highlevelmonitorconfigurationapi.h>

// load the relevant libary
#pragma comment(lib, "Dxva2.lib")
#pragma comment(lib, "User32.lib")


ddcmonitors::ddcmonitors()
{
	// get the list of monitors
	EnumDisplayMonitors(NULL, NULL, ddcmonitors::MonitorEnumProc, reinterpret_cast<LPARAM>(&MonitorHandles));
}


ddcmonitors::~ddcmonitors()
{
	// release the monitor handle
	for (auto hMonitor : MonitorHandles)
	{
		DestroyPhysicalMonitor(hMonitor);
	}
}

// A static callback used for retrieving the list of monitors
// Called by the EnumDisplayMonitors() method used in Monitors class constructor (once for each monitor)
BOOL CALLBACK ddcmonitors::MonitorEnumProc(HMONITOR hMonitor, HDC hdcMonitor,
	LPRECT lprcMonitor, LPARAM dwData)
{
	std::set<HANDLE>* pMonitors = reinterpret_cast<std::set<HANDLE>*>(dwData);
	DWORD nMonitorCount;

	if (GetNumberOfPhysicalMonitorsFromHMONITOR(hMonitor, &nMonitorCount))
	{
		PHYSICAL_MONITOR* pMons = new PHYSICAL_MONITOR[nMonitorCount];

		if (GetPhysicalMonitorsFromHMONITOR(hMonitor, nMonitorCount, pMons))
		{
			for (DWORD i = 0; i < nMonitorCount; i++)
			{
				pMonitors->insert(pMons[i].hPhysicalMonitor);
			}
		}
		delete[] pMons;
	}

	return TRUE;
}

// Get number of displays
// param:
//		none
int ddcmonitors::GetAmount()
{
	return (int)MonitorHandles.size();
}

// Set power mode of the monitor using DDC/CI
// param:
//		nPower	Power::<On, Standby, Suspend, Off, HardOff>
void ddcmonitors::SetPower(Power nPower)
{
	const BYTE PowerMode = 0xD6;

	if ((nPower < 1) || (nPower > 5))
	{
		return;
	}

	for (auto hMonitor : MonitorHandles)
	{
		SetVCPFeature(hMonitor, PowerMode, nPower);
	}
}

// Set backlight of the monitor using DDC/CI
// param:
//		brightness <0 ... 100>
void ddcmonitors::SetBackLight(int brightness)
{
	if (brightness > 100)
	{
		brightness = 100;
	}
	int setbrightness = brightness;
	int minbrightness, currentbrightness, maxbrightness;
	for (auto hMonitor : MonitorHandles)
	{
		GetMonitorBrightness(hMonitor, (LPDWORD)(&minbrightness), (LPDWORD)(&currentbrightness), (LPDWORD)(&maxbrightness));
		setbrightness = ((brightness * (maxbrightness - minbrightness)) / 100) + minbrightness;
		if (setbrightness != currentbrightness)
		{
			SetMonitorBrightness(hMonitor, (DWORD)setbrightness);
		}
	}
}
