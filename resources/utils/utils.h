/**
 *
 * (c) Copyright Ascensio System SIA 2024
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 */

// convenient macro definitions
#if defined(__linux__) || defined(__linux)
#define _LINUX
#elif defined(__APPLE__) || defined(__MACH__)
#define _MAC
#endif

#include <string>

#ifdef _WIN32
#include <wchar.h>
#include <windows.h>
#endif

#if defined(_LINUX) || defined(_MAC)
#include <unistd.h>
#include <string.h>
#include <sys/stat.h>
#endif

#ifdef _MAC
#include <mach-o/dyld.h>
#endif

#ifdef CreateFile
#undef CreateFile
#endif

namespace NSUtils
{
#ifdef _WIN32
	#define FILE_SEPARATOR '\\'
	inline void WriteCodepoint(int code, wchar_t* unicodes_cur)
	{
		if (code < 0x10000)
		{
			*unicodes_cur++ = code;
		}
		else
		{
			code -= 0x10000;
			*unicodes_cur++ = 0xD800 | ((code >> 10) & 0x03FF);
			*unicodes_cur++ = 0xDC00 | (code & 0x03FF);
		}
	}
#else
	#define FILE_SEPARATOR '/'
	inline void WriteCodepoint(int code, wchar_t* unicodes_cur)
	{
		*unicodes_cur++ = (wchar_t)code;
	}
#endif

	std::wstring GetStringFromUtf8(const unsigned char* utf8, size_t length)
	{
		wchar_t* unicodes = new wchar_t[length + 1];
		wchar_t* unicodes_cur = unicodes;
		size_t index = 0;

		while (index < length)
		{
			unsigned char byteMain = utf8[index];
			if (0x00 == (byteMain & 0x80))
			{
				// 1 byte
				WriteCodepoint(byteMain, unicodes_cur);
				++index;
			}
			else if (0x00 == (byteMain & 0x20))
			{
				// 2 byte
				int val = 0;
				if ((index + 1) < length)
				{
					val = (int)(((byteMain & 0x1F) << 6) |
								(utf8[index + 1] & 0x3F));
				}

				WriteCodepoint(val, unicodes_cur);
				index += 2;
			}
			else if (0x00 == (byteMain & 0x10))
			{
				// 3 byte
				int val = 0;
				if ((index + 2) < length)
				{
					val = (int)(((byteMain & 0x0F) << 12) |
								((utf8[index + 1] & 0x3F) << 6) |
								(utf8[index + 2] & 0x3F));
				}

				WriteCodepoint(val, unicodes_cur);
				index += 3;
			}
			else if (0x00 == (byteMain & 0x0F))
			{
				// 4 byte
				int val = 0;
				if ((index + 3) < length)
				{
					val = (int)(((byteMain & 0x07) << 18) |
								((utf8[index + 1] & 0x3F) << 12) |
								((utf8[index + 2] & 0x3F) << 6) |
								(utf8[index + 3] & 0x3F));
				}

				WriteCodepoint(val, unicodes_cur);
				index += 4;
			}
			else if (0x00 == (byteMain & 0x08))
			{
				// 4 byte
				int val = 0;
				if ((index + 3) < length)
				{
					val = (int)(((byteMain & 0x07) << 18) |
								((utf8[index + 1] & 0x3F) << 12) |
								((utf8[index + 2] & 0x3F) << 6) |
								(utf8[index + 3] & 0x3F));
				}

				WriteCodepoint(val, unicodes_cur);
				index += 4;
			}
			else if (0x00 == (byteMain & 0x04))
			{
				// 5 byte
				int val = 0;
				if ((index + 4) < length)
				{
					val = (int)(((byteMain & 0x03) << 24) |
								((utf8[index + 1] & 0x3F) << 18) |
								((utf8[index + 2] & 0x3F) << 12) |
								((utf8[index + 3] & 0x3F) << 6) |
								(utf8[index + 4] & 0x3F));
				}

				WriteCodepoint(val, unicodes_cur);
				index += 5;
			}
			else
			{
				// 6 byte
				int val = 0;
				if ((index + 5) < length)
				{
					val = (int)(((byteMain & 0x01) << 30) |
								((utf8[index + 1] & 0x3F) << 24) |
								((utf8[index + 2] & 0x3F) << 18) |
								((utf8[index + 3] & 0x3F) << 12) |
								((utf8[index + 4] & 0x3F) << 6) |
								(utf8[index + 5] & 0x3F));
				}

				WriteCodepoint(val, unicodes_cur);
				index += 5;
			}
		}

		*unicodes_cur++ = 0;

		std::wstring sOutput(unicodes);

		delete[] unicodes;

		return sOutput;
	}

	void GetUtf8StringFromUnicode_4bytes(const wchar_t* pUnicodes, LONG lCount, BYTE*& pData, LONG& lOutputCount, bool bIsBOM)
	{
		if (NULL == pData)
		{
			pData = new BYTE[6 * lCount + 3 + 1];
		}

		BYTE* pCodesCur = pData;
		if (bIsBOM)
		{
			pCodesCur[0] = 0xEF;
			pCodesCur[1] = 0xBB;
			pCodesCur[2] = 0xBF;
			pCodesCur += 3;
		}

		const wchar_t* pEnd = pUnicodes + lCount;
		const wchar_t* pCur = pUnicodes;

		while (pCur < pEnd)
		{
			unsigned int code = (unsigned int)*pCur++;

			if (code < 0x80)
			{
				*pCodesCur++ = (BYTE)code;
			}
			else if (code < 0x0800)
			{
				*pCodesCur++ = 0xC0 | (code >> 6);
				*pCodesCur++ = 0x80 | (code & 0x3F);
			}
			else if (code < 0x10000)
			{
				*pCodesCur++ = 0xE0 | (code >> 12);
				*pCodesCur++ = 0x80 | (code >> 6 & 0x3F);
				*pCodesCur++ = 0x80 | (code & 0x3F);
			}
			else if (code < 0x1FFFFF)
			{
				*pCodesCur++ = 0xF0 | (code >> 18);
				*pCodesCur++ = 0x80 | (code >> 12 & 0x3F);
				*pCodesCur++ = 0x80 | (code >> 6 & 0x3F);
				*pCodesCur++ = 0x80 | (code & 0x3F);
			}
			else if (code < 0x3FFFFFF)
			{
				*pCodesCur++ = 0xF8 | (code >> 24);
				*pCodesCur++ = 0x80 | (code >> 18 & 0x3F);
				*pCodesCur++ = 0x80 | (code >> 12 & 0x3F);
				*pCodesCur++ = 0x80 | (code >> 6 & 0x3F);
				*pCodesCur++ = 0x80 | (code & 0x3F);
			}
			else if (code < 0x7FFFFFFF)
			{
				*pCodesCur++ = 0xFC | (code >> 30);
				*pCodesCur++ = 0x80 | (code >> 24 & 0x3F);
				*pCodesCur++ = 0x80 | (code >> 18 & 0x3F);
				*pCodesCur++ = 0x80 | (code >> 12 & 0x3F);
				*pCodesCur++ = 0x80 | (code >> 6 & 0x3F);
				*pCodesCur++ = 0x80 | (code & 0x3F);
			}
		}

		lOutputCount = (LONG)(pCodesCur - pData);
		*pCodesCur++ = 0;
	}

	void GetUtf8StringFromUnicode_2bytes(const wchar_t* pUnicodes, LONG lCount, BYTE*& pData, LONG& lOutputCount, bool bIsBOM)
	{
		if (NULL == pData)
		{
			pData = new BYTE[6 * lCount + 3 + 1];
		}

		BYTE* pCodesCur = pData;
		if (bIsBOM)
		{
			pCodesCur[0] = 0xEF;
			pCodesCur[1] = 0xBB;
			pCodesCur[2] = 0xBF;
			pCodesCur += 3;
		}

		const wchar_t* pEnd = pUnicodes + lCount;
		const wchar_t* pCur = pUnicodes;

		while (pCur < pEnd)
		{
			unsigned int code = (unsigned int)*pCur++;
			if (code >= 0xD800 && code <= 0xDFFF && pCur < pEnd)
			{
				code = 0x10000 + (((code & 0x3FF) << 10) | (0x03FF & *pCur++));
			}

			if (code < 0x80)
			{
				*pCodesCur++ = (BYTE)code;
			}
			else if (code < 0x0800)
			{
				*pCodesCur++ = 0xC0 | (code >> 6);
				*pCodesCur++ = 0x80 | (code & 0x3F);
			}
			else if (code < 0x10000)
			{
				*pCodesCur++ = 0xE0 | (code >> 12);
				*pCodesCur++ = 0x80 | ((code >> 6) & 0x3F);
				*pCodesCur++ = 0x80 | (code & 0x3F);
			}
			else if (code < 0x1FFFFF)
			{
				*pCodesCur++ = 0xF0 | (code >> 18);
				*pCodesCur++ = 0x80 | ((code >> 12) & 0x3F);
				*pCodesCur++ = 0x80 | ((code >> 6) & 0x3F);
				*pCodesCur++ = 0x80 | (code & 0x3F);
			}
			else if (code < 0x3FFFFFF)
			{
				*pCodesCur++ = 0xF8 | (code >> 24);
				*pCodesCur++ = 0x80 | ((code >> 18) & 0x3F);
				*pCodesCur++ = 0x80 | ((code >> 12) & 0x3F);
				*pCodesCur++ = 0x80 | ((code >> 6) & 0x3F);
				*pCodesCur++ = 0x80 | (code & 0x3F);
			}
			else if (code < 0x7FFFFFFF)
			{
				*pCodesCur++ = 0xFC | (code >> 30);
				*pCodesCur++ = 0x80 | ((code >> 24) & 0x3F);
				*pCodesCur++ = 0x80 | ((code >> 18) & 0x3F);
				*pCodesCur++ = 0x80 | ((code >> 12) & 0x3F);
				*pCodesCur++ = 0x80 | ((code >> 6) & 0x3F);
				*pCodesCur++ = 0x80 | (code & 0x3F);
			}
		}

		lOutputCount = (LONG)(pCodesCur - pData);
		*pCodesCur++ = 0;
	}

	void GetUtf8StringFromUnicode(const wchar_t* pUnicodes, LONG lCount, BYTE*& pData, LONG& lOutputCount)
	{
		if (NULL == pUnicodes || 0 == lCount)
		{
			pData = NULL;
			lOutputCount = 0;
			return;
		}

		if (sizeof(WCHAR) == 2)
			return GetUtf8StringFromUnicode_2bytes(pUnicodes, lCount, pData, lOutputCount, false);
		return GetUtf8StringFromUnicode_4bytes(pUnicodes, lCount, pData, lOutputCount, false);
	}

	std::string GetUtf8StringFromUnicode(const wchar_t* pUnicodes, LONG lCount)
	{
		if (NULL == pUnicodes || 0 == lCount)
			return "";

		BYTE* pData = NULL;
		LONG lLen = 0;

		GetUtf8StringFromUnicode(pUnicodes, lCount, pData, lLen);

		std::string s((char*)pData, lLen);

		delete[] pData;
		return s;
	}
}

#define U_TO_UTF8(val) NSUtils::GetUtf8StringFromUnicode(val.c_str(), (LONG)val.length())

namespace NSUtils
{
	#define NS_FILE_MAX_PATH 32768
	std::wstring GetProcessPath()
	{
#ifdef _WIN32
		wchar_t buf [NS_FILE_MAX_PATH];
		GetModuleFileNameW(GetModuleHandle(NULL), buf, NS_FILE_MAX_PATH);
		return std::wstring(buf);
#endif

#if defined(_LINUX) || defined(_MAC)
		char buf[NS_FILE_MAX_PATH];
		memset(buf, 0, NS_FILE_MAX_PATH);
		if (readlink ("/proc/self/exe", buf, NS_FILE_MAX_PATH) <= 0)
		{
#ifdef _MAC
			uint32_t _size = NS_FILE_MAX_PATH;
			_NSGetExecutablePath(buf, &_size);
#endif
		}
		return GetStringFromUtf8((unsigned char*)buf, strlen(buf));
#endif
	}

	std::wstring GetProcessDirectory()
	{
		std::wstring path = GetProcessPath();
		size_t pos = path.find_last_of(FILE_SEPARATOR);
		if (pos != std::wstring::npos)
			path = path.substr(0, pos);
		return path;
	}

	std::wstring GetResourcesDirectory()
	{
		std::wstring path = GetProcessDirectory();
		while (!path.empty())
		{
			size_t pos = path.find_last_of(FILE_SEPARATOR);
			if (pos != std::wstring::npos)
			{
				std::wstring currDir = path.substr(pos + 1);
				// update to parent directory
				path = path.substr(0, pos);
				if (currDir == L"out")
				{
					path += FILE_SEPARATOR;
					path += L"resources";
					break;
				}
			}
			else
			{
				path = L"";
			}
		}

		return path;
	}
}
