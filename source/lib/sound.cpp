/*
AutoHotkey

Copyright 2003-2009 Chris Mallett (support@autohotkey.com)

This program is free software; you can redistribute it and/or
modify it under the terms of the GNU General Public License
as published by the Free Software Foundation; either version 2
of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.
*/

#include "stdafx.h" // pre-compiled headers
#include "script.h"
#include "globaldata.h"
#include "application.h"

#include "qmath.h"
#include <mmdeviceapi.h>
#include <endpointvolume.h>
#include <functiondiscoverykeys.h>

#include "script_func_impl.h"



#pragma region Sound support functions - Vista and later

enum class SoundControlType
{
	Volume,
	Mute,
	Name,
	IID
};

LPWSTR SoundDeviceGetName(IMMDevice *aDev)
{
	IPropertyStore *store;
	PROPVARIANT prop;
	if (SUCCEEDED(aDev->OpenPropertyStore(STGM_READ, &store)))
	{
		if (FAILED(store->GetValue(PKEY_Device_FriendlyName, &prop)))
			prop.pwszVal = nullptr;
		store->Release();
		return prop.pwszVal;
	}
	return nullptr;
}


HRESULT SoundSetGet_GetDevice(LPTSTR aDeviceString, IMMDevice *&aDevice)
{
	IMMDeviceEnumerator *deviceEnum;
	IMMDeviceCollection *devices;
	HRESULT hr;

	aDevice = NULL;

	hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void**)&deviceEnum);
	if (SUCCEEDED(hr))
	{
		if (!*aDeviceString)
		{
			// Get default playback device.
			hr = deviceEnum->GetDefaultAudioEndpoint(eRender, eConsole, &aDevice);
		}
		else
		{
			// Parse Name:Index, Name or Index.
			int target_index = 0;
			LPWSTR target_name = L"";
			size_t target_name_length = 0;
			LPTSTR delim = _tcsrchr(aDeviceString, ':');
			if (delim)
			{
				target_index = ATOI(delim + 1) - 1;
				target_name = aDeviceString;
				target_name_length = delim - aDeviceString;
			}
			else if (IsNumeric(aDeviceString, FALSE, FALSE))
			{
				target_index = ATOI(aDeviceString) - 1;
			}
			else
			{
				target_name = aDeviceString;
				target_name_length = _tcslen(aDeviceString);
			}

			// Enumerate devices; include unplugged devices so that indices don't change when a device is plugged in.
			hr = deviceEnum->EnumAudioEndpoints(eAll, DEVICE_STATE_ACTIVE | DEVICE_STATE_UNPLUGGED, &devices);
			if (SUCCEEDED(hr))
			{
				if (target_name_length)
				{
					IMMDevice *dev;
					for (UINT u = 0; SUCCEEDED(hr = devices->Item(u, &dev)); ++u)
					{
						if (LPWSTR dev_name = SoundDeviceGetName(dev))
						{
							if (!_wcsnicmp(dev_name, target_name, target_name_length)
								&& target_index-- == 0)
							{
								CoTaskMemFree(dev_name);
								aDevice = dev;
								break;
							}
							CoTaskMemFree(dev_name);
						}
						dev->Release();
					}
				}
				else
					hr = devices->Item((UINT)target_index, &aDevice);
				devices->Release();
			}
		}
		deviceEnum->Release();
	}
	return hr;
}


struct SoundComponentSearch
{
	// Parameters of search:
	GUID target_iid;
	int target_instance;
	SoundControlType target_control;
	TCHAR target_name[128]; // Arbitrary; probably larger than any real component name.
	// Internal use/results:
	IUnknown *control;
	LPWSTR name; // Valid only when target_control == SoundControlType::Name.
	int count;
	// Internal use:
	DataFlow data_flow;
	bool ignore_remaining_subunits;
};


void SoundConvertComponent(LPTSTR aBuf, SoundComponentSearch &aSearch)
{
	if (IsNumeric(aBuf) == PURE_INTEGER)
	{
		*aSearch.target_name = '\0';
		aSearch.target_instance = ATOI(aBuf);
	}
	else
	{
		tcslcpy(aSearch.target_name, aBuf, _countof(aSearch.target_name));
		LPTSTR colon_pos = _tcsrchr(aSearch.target_name, ':');
		aSearch.target_instance = colon_pos ? ATOI(colon_pos + 1) : 1;
		if (colon_pos)
			*colon_pos = '\0';
	}
}


bool SoundSetGet_FindComponent(IPart *aRoot, SoundComponentSearch &aSearch)
{
	HRESULT hr;
	IPartsList *parts;
	IPart *part;
	UINT part_count;
	PartType part_type;
	LPWSTR part_name;
	
	bool check_name = *aSearch.target_name;

	if (aSearch.data_flow == In)
		hr = aRoot->EnumPartsIncoming(&parts);
	else
		hr = aRoot->EnumPartsOutgoing(&parts);
	
	if (FAILED(hr))
		return NULL;

	if (FAILED(parts->GetCount(&part_count)))
		part_count = 0;
	
	for (UINT i = 0; i < part_count; ++i)
	{
		if (FAILED(parts->GetPart(i, &part)))
			continue;

		if (SUCCEEDED(part->GetPartType(&part_type)))
		{
			if (part_type == Connector)
			{
				if (   part_count == 1 // Ignore Connectors with no Subunits of their own.
					&& (!check_name || (SUCCEEDED(part->GetName(&part_name)) && !_wcsicmp(part_name, aSearch.target_name)))   )
				{
					if (++aSearch.count == aSearch.target_instance)
					{
						switch (aSearch.target_control)
						{
						case SoundControlType::IID:
							// Permit retrieving the IPart or IConnector itself.  Since there may be
							// multiple connected Subunits (and they can be enumerated or retrieved
							// via the Connector IPart), this is only done for the Connector.
							part->QueryInterface(aSearch.target_iid, (void **)&aSearch.control);
							break;
						case SoundControlType::Name:
							// check_name would typically be false in this case, since a value of true
							// would mean the caller already knows the component's name.
							part->GetName(&aSearch.name);
							break;
						}
						part->Release();
						parts->Release();
						return true;
					}
				}
			}
			else // Subunit
			{
				// Recursively find the Connector nodes linked to this part.
				if (SoundSetGet_FindComponent(part, aSearch))
				{
					// A matching connector part has been found with this part as one of the nodes used
					// to reach it.  Therefore, if this part supports the requested control interface,
					// it can in theory be used to control the component.  An example path might be:
					//    Output < Master Mute < Master Volume < Sum < Mute < Volume < CD Audio
					// Parts are considered from right to left, as we return from recursion.
					if (!aSearch.control && !aSearch.ignore_remaining_subunits)
					{
						// Query this part for the requested interface and let caller check the result.
						part->Activate(CLSCTX_ALL, aSearch.target_iid, (void **)&aSearch.control);

						// If this subunit has siblings, ignore any controls further up the line
						// as they're likely shared by other components (i.e. master controls).
						if (part_count > 1)
							aSearch.ignore_remaining_subunits = true;
					}
					part->Release();
					parts->Release();
					return true;
				}
			}
		}

		part->Release();
	}

	parts->Release();
	return false;
}


bool SoundSetGet_FindComponent(IMMDevice *aDevice, SoundComponentSearch &aSearch)
{
	IDeviceTopology *topo;
	IConnector *conn, *conn_to;
	IPart *part;

	aSearch.count = 0;
	aSearch.control = nullptr;
	aSearch.name = nullptr;
	aSearch.ignore_remaining_subunits = false;
	
	if (SUCCEEDED(aDevice->Activate(__uuidof(IDeviceTopology), CLSCTX_ALL, NULL, (void**)&topo)))
	{
		if (SUCCEEDED(topo->GetConnector(0, &conn)))
		{
			if (SUCCEEDED(conn->GetDataFlow(&aSearch.data_flow))
			 && SUCCEEDED(conn->GetConnectedTo(&conn_to)))
			{
				if (SUCCEEDED(conn_to->QueryInterface(&part)))
				{
					// Search; the result is stored in the search struct.
					SoundSetGet_FindComponent(part, aSearch);
					part->Release();
				}
				conn_to->Release();
			}
			conn->Release();
		}
		topo->Release();
	}

	return aSearch.count == aSearch.target_instance;
}

#pragma endregion



BIF_DECL(BIF_Sound)
{
	LPTSTR aSetting = nullptr;
	SoundComponentSearch search;
	if (_f_callee_id >= FID_SoundSetVolume)
	{
		search.target_control = SoundControlType(_f_callee_id - FID_SoundSetVolume);
		aSetting = ParamIndexToString(0, _f_number_buf);
		++aParam;
		--aParamCount;
	}
	else
		search.target_control = SoundControlType(_f_callee_id);

	switch (search.target_control)
	{
	case SoundControlType::Volume:
		search.target_iid = __uuidof(IAudioVolumeLevel);
		break;
	case SoundControlType::Mute:
		search.target_iid = __uuidof(IAudioMute);
		break;
	case SoundControlType::IID:
		LPTSTR iid; iid = ParamIndexToString(0);
		if (*iid != '{' || FAILED(CLSIDFromString(iid, &search.target_iid)))
			_f_throw_param(0);
		++aParam;
		--aParamCount;
		break;
	}

	_f_param_string_opt(aComponent, 0);
	_f_param_string_opt(aDevice, 1);

#define SOUND_MODE_IS_SET aSetting // Boolean: i.e. if it's not NULL, the mode is "SET".
	_f_set_retval_p(_T(""), 0); // Set default.

	float setting_scalar;
	if (SOUND_MODE_IS_SET)
	{
		setting_scalar = (float)(ATOF(aSetting) / 100);
		if (setting_scalar < -1)
			setting_scalar = -1;
		else if (setting_scalar > 1)
			setting_scalar = 1;
	}

	// Does user want to adjust the current setting by a certain amount?
	bool adjust_current_setting = aSetting && (*aSetting == '-' || *aSetting == '+');

	IMMDevice *device;
	HRESULT hr;

	hr = SoundSetGet_GetDevice(aDevice, device);
	if (FAILED(hr))
		_f_throw(ERR_SOUND_DEVICE, ErrorPrototype::Target);

	LPCTSTR errorlevel = NULL;
	float result_float;
	BOOL result_bool;

	if (!*aComponent) // Component is Master (omitted).
	{
		if (search.target_control == SoundControlType::IID)
		{
			void *result_ptr = nullptr;
			if (FAILED(device->QueryInterface(search.target_iid, (void **)&result_ptr)))
				device->Activate(search.target_iid, CLSCTX_ALL, NULL, &result_ptr);
			// For consistency with ComObjQuery, the result is returned even on failure.
			aResultToken.SetValue((UINT_PTR)result_ptr);
		}
		else if (search.target_control == SoundControlType::Name)
		{
			if (auto name = SoundDeviceGetName(device))
			{
				aResultToken.Return(name);
				CoTaskMemFree(name);
			}
		}
		else
		{
			// For Master/Speakers, use the IAudioEndpointVolume interface.  Some devices support master
			// volume control, but do not actually have a volume subunit (so the other method would fail).
			IAudioEndpointVolume *aev;
			hr = device->Activate(__uuidof(IAudioEndpointVolume), CLSCTX_ALL, NULL, (void**)&aev);
			if (SUCCEEDED(hr))
			{
				if (search.target_control == SoundControlType::Volume)
				{
					if (!SOUND_MODE_IS_SET || adjust_current_setting)
					{
						hr = aev->GetMasterVolumeLevelScalar(&result_float);
					}

					if (SUCCEEDED(hr))
					{
						if (SOUND_MODE_IS_SET)
						{
							if (adjust_current_setting)
							{
								setting_scalar += result_float;
								if (setting_scalar > 1)
									setting_scalar = 1;
								else if (setting_scalar < 0)
									setting_scalar = 0;
							}
							hr = aev->SetMasterVolumeLevelScalar(setting_scalar, NULL);
						}
						else
						{
							result_float *= 100.0;
						} 
					}
				}
				else // Mute.
				{
					if (!SOUND_MODE_IS_SET || adjust_current_setting)
					{
						hr = aev->GetMute(&result_bool);
					}
					if (SOUND_MODE_IS_SET && SUCCEEDED(hr))
					{
						hr = aev->SetMute(adjust_current_setting ? !result_bool : setting_scalar > 0, NULL);
					}
				}
				aev->Release();
			}
		}
	}
	else
	{
		SoundConvertComponent(aComponent, search);
		
		if (!SoundSetGet_FindComponent(device, search))
		{
			errorlevel = ERR_SOUND_COMPONENT;
		}
		else if (search.target_control == SoundControlType::IID)
		{
			aResultToken.SetValue((UINT_PTR)search.control);
			search.control = nullptr; // Don't release it.
		}
		else if (search.target_control == SoundControlType::Name)
		{
			aResultToken.Return(search.name);
		}
		else if (!search.control)
		{
			errorlevel = ERR_SOUND_CONTROLTYPE;
		}
		else if (search.target_control == SoundControlType::Volume)
		{
			IAudioVolumeLevel *avl = (IAudioVolumeLevel *)search.control;
				
			UINT channel_count = 0;
			if (FAILED(avl->GetChannelCount(&channel_count)))
				goto control_fail;
				
			float *level = (float *)_alloca(sizeof(float) * 3 * channel_count);
			float *level_min = level + channel_count;
			float *level_range = level_min + channel_count;
			float f, db, min_db, max_db, max_level = 0;

			for (UINT i = 0; i < channel_count; ++i)
			{
				if (FAILED(avl->GetLevel(i, &db)) ||
					FAILED(avl->GetLevelRange(i, &min_db, &max_db, &f)))
					goto control_fail;
				// Convert dB to scalar.
				level_min[i] = (float)qmathExp10(min_db/20);
				level_range[i] = (float)qmathExp10(max_db/20) - level_min[i];
				// Compensate for differing level ranges. (No effect if range is -96..0 dB.)
				level[i] = ((float)qmathExp10(db/20) - level_min[i]) / level_range[i];
				// Windows reports the highest level as the overall volume.
				if (max_level < level[i])
					max_level = level[i];
			}

			if (SOUND_MODE_IS_SET)
			{
				if (adjust_current_setting)
				{
					setting_scalar += max_level;
					if (setting_scalar > 1)
						setting_scalar = 1;
					else if (setting_scalar < 0)
						setting_scalar = 0;
				}

				for (UINT i = 0; i < channel_count; ++i)
				{
					f = setting_scalar;
					if (max_level)
						f *= (level[i] / max_level); // Preserve balance.
					// Compensate for differing level ranges.
					f = level_min[i] + f * level_range[i];
					// Convert scalar to dB.
					level[i] = 20 * (float)qmathLog10(f);
				}

				hr = avl->SetLevelAllChannels(level, channel_count, NULL);
			}
			else
			{
				result_float = max_level * 100;
			}
		}
		else if (search.target_control == SoundControlType::Mute)
		{
			IAudioMute *am = (IAudioMute *)search.control;

			if (!SOUND_MODE_IS_SET || adjust_current_setting)
			{
				hr = am->GetMute(&result_bool);
			}
			if (SOUND_MODE_IS_SET && SUCCEEDED(hr))
			{
				hr = am->SetMute(adjust_current_setting ? !result_bool : setting_scalar > 0, NULL);
			}
		}
control_fail:
		if (search.control)
			search.control->Release();
		if (search.name)
			CoTaskMemFree(search.name);
	}

	device->Release();
	
	if (errorlevel)
		_f_throw(errorlevel, ErrorPrototype::Target);
	if (FAILED(hr)) // This would be rare and unexpected.
		_f_throw_win32(hr);

	if (SOUND_MODE_IS_SET)
		return;

	switch (search.target_control)
	{
	case SoundControlType::Volume: aResultToken.Return(result_float); break;
	case SoundControlType::Mute: aResultToken.Return(result_bool); break;
	}
}



bif_impl FResult SoundPlay(LPCTSTR aFilespec, LPCTSTR aWait)
{
	auto cp = omit_leading_whitespace(aFilespec);
	if (*cp == '*')
	{
		// ATOU() returns 0xFFFFFFFF for -1, which is relied upon to support the -1 sound.
		// Even if there are no enabled sound devices, MessageBeep() indicates success
		// (tested on Windows 10).  It's hard to imagine how the script would handle a
		// failure anyway, so just omit error checking.
		MessageBeep(ATOU(cp + 1));
		return OK;
	}
	// See http://msdn.microsoft.com/library/default.asp?url=/library/en-us/multimed/htm/_win32_play.asp
	// for some documentation mciSendString() and related.
	// MAX_PATH note: There's no chance this API supports long paths even on Windows 10.
	// Research indicates it limits paths to 127 chars (not even MAX_PATH), but since there's
	// no apparent benefit to reducing it, we'll keep this size to ensure backward-compatibility.
	// Note that using a relative path does not help, but using short (8.3) names does.
	TCHAR buf[MAX_PATH * 2]; // Allow room for filename and commands.
	mciSendString(_T("status ") SOUNDPLAY_ALIAS _T(" mode"), buf, _countof(buf), NULL);
	if (*buf) // "playing" or "stopped" (so close it before trying to re-open with a new aFilespec).
		mciSendString(_T("close ") SOUNDPLAY_ALIAS, NULL, 0, NULL);
	sntprintf(buf, _countof(buf), _T("open \"%s\" alias ") SOUNDPLAY_ALIAS, aFilespec);
	if (mciSendString(buf, NULL, 0, NULL)) // Failure.
		return FR_E_FAILED;
	g_SoundWasPlayed = true;  // For use by Script's destructor.
	if (mciSendString(_T("play ") SOUNDPLAY_ALIAS, NULL, 0, NULL)) // Failure.
		return FR_E_FAILED;
	// Otherwise, the sound is now playing.
	if (  !(aWait && (aWait[0] == '1' && !aWait[1] || !_tcsicmp(aWait, _T("Wait"))))  )
		return OK;
	// Otherwise, caller wants us to wait until the file is done playing.  To allow our app to remain
	// responsive during this time, use a loop that checks our message queue:
	// Older method: "mciSendString("play " SOUNDPLAY_ALIAS " wait", NULL, 0, NULL)"
	for (;;)
	{
		mciSendString(_T("status ") SOUNDPLAY_ALIAS _T(" mode"), buf, _countof(buf), NULL);
		if (!*buf) // Probably can't happen given the state we're in.
			break;
		if (!_tcscmp(buf, _T("stopped"))) // The sound is done playing.
		{
			mciSendString(_T("close ") SOUNDPLAY_ALIAS, NULL, 0, NULL);
			break;
		}
		// Sleep a little longer than normal because I'm not sure how much overhead
		// and CPU utilization the above incurs:
		MsgSleep(20);
	}
	return OK;
}



bif_impl void SoundBeep(int *aFrequency, int *aDuration)
{
	Beep(aFrequency ? *aFrequency : 523, aDuration && *aDuration >= 0 ? *aDuration : 150);
}
