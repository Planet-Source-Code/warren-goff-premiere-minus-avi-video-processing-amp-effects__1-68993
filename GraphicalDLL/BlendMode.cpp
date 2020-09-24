#include "stdafx.h"
#include <windows.h>
#include <math.h>
#include "BlendMode.h"

// Average blend mode
inline BYTE AverageMode (BYTE Color1, BYTE Color2)
{
	return ((Color1 + Color2) >> 1);
}

// Multiply blend mode
inline BYTE MultiplyMode (BYTE Color1, BYTE Color2)
{
	return ((Color1 * Color2) >> 8);
}

// Screen blend mode
inline BYTE ScreenMode (BYTE Color1, BYTE Color2)
{
	return (255 - ((255 - Color1) * (255 - Color2) >> 8));
}

// Darken blend mode
inline BYTE DarkenMode (BYTE Color1, BYTE Color2)
{
	return ((Color1 < Color2) ? Color1 : Color2);
}

// Lighten blend mode
inline BYTE LightenMode (BYTE Color1, BYTE Color2)
{
	return ((Color1 > Color2) ? Color1 : Color2);
}

// Difference blend mode
inline BYTE DifferenceMode (BYTE Color1, BYTE Color2)
{
	return (abs (Color1 - Color2));
}

// Negation blend mode
inline BYTE NegationMode (BYTE Color1, BYTE Color2)
{
	return (255 - abs (Color1 - Color2));
}

// Exclusion blend mode
inline BYTE ExclusionMode (BYTE Color1, BYTE Color2)
{
	return (Color1 + Color2 - ((Color1 * Color2) >> 7));
}

// Overlay blend mode
inline BYTE OverlayMode (BYTE Color1, BYTE Color2)
{
	if (Color1 < 128)
		return ((Color1 * Color2) >> 7);
	else
		return (255 - ((255 - Color1) * (255 - Color2) >> 7));
}

// Hard light blend mode
inline BYTE HardLightMode (BYTE Color1, BYTE Color2)
{
	if (Color2 < 128)
		return ((Color1 * Color2) >> 7);
	else
		return (255 - ((255 - Color1) * (255 - Color2) >> 7));
}

// Soft light blend mode
inline BYTE SoftLightMode (BYTE Color1, BYTE Color2)
{
	BYTE lTemp = (Color1 * Color2) >> 8;

	return (lTemp + (Color1 * (255 - ((255 - Color1) * (255 - Color2) >> 8) - lTemp) >> 8));
}

// Color dodge blend mode
inline BYTE ColorDodgeMode (BYTE Color1, BYTE Color2)
{
	int lTemp;

	if (Color2 == 255)
		return (255);
	else
	{
		lTemp = (Color1 << 8) / (255 - Color2);
		return ((lTemp > 255) ? 255 : (BYTE)lTemp);
	}
}

// Color burn blend mode
inline BYTE ColorBurnMode (BYTE Color1, BYTE Color2)
{
	int lTemp;

	if (Color2 == 0)
		return (0);
	else
	{
		lTemp = 255 - (((255 - Color1) << 8) / Color2);
		return ((lTemp < 0) ? 0 : (BYTE)lTemp);
	}
}

// Soft dodge blend mode
inline BYTE SoftDodgeMode (BYTE Color1, BYTE Color2)
{
	int lTemp;

	if (Color1 + Color2 < 256)
	{
		if (Color2 == 255)
			return (255);
		else
		{
			lTemp = (Color1 << 7) / (255 - Color2);
			return ((lTemp > 255) ? 255 : (BYTE)lTemp);
		}
	}
	else
	{
		if (Color2 == 255)
			return (255);
		else
		{
			lTemp = 255 - (((255 - Color2) << 7) / Color1);
			return ((lTemp < 0) ? 0 : (BYTE)lTemp);
		}
	}
}

// Soft burn blend mode
inline BYTE SoftBurnMode (BYTE Color1, BYTE Color2)
{
	int lTemp;

	if (Color1 + Color2 < 256)
	{
		if (Color1 == 255)
			return (255);
		else
		{
			lTemp = (Color2 << 7) / (255 - Color1);
			return ((lTemp > 255) ? 255 : (BYTE)lTemp);
		}
	}
	else
	{
		if (Color1 == 255)
			return (255);
		else
		{
			lTemp = 255 - (((255 - Color1) << 7) / Color2);
			return ((lTemp < 0) ? 0 : (BYTE)lTemp);
		}
	}
}

// Reflect blend mode
inline BYTE ReflectMode (BYTE Color1, BYTE Color2)
{
	int lTemp;

	if (Color2 == 255)
		return (255);
	else
	{
		lTemp = (Color1 * Color1) / (255 - Color2);
		return ((lTemp > 255) ? 255 : (BYTE)lTemp);
	}
}

// Glow blend mode
inline BYTE GlowMode (BYTE Color1, BYTE Color2)
{
	int lTemp;

	if (Color1 == 255)
		return (255);
	else
	{
		lTemp = (Color2 * Color2) / (255 - Color1);
		return ((lTemp > 255) ? 255 : (BYTE)lTemp);
	}
}

// Freeze blend mode
inline BYTE FreezeMode (BYTE Color1, BYTE Color2)
{
	int lTemp;

	if (Color2 == 0)
		return (0);
	else
	{
		lTemp = 255 - ((255 - Color1) * (255 - Color1)) / Color2;
		return ((lTemp < 0) ? 0 : (BYTE)lTemp);
	}
}

// Heat blend mode
inline BYTE HeatMode (BYTE Color1, BYTE Color2)
{
	int lTemp;

	if (Color1 == 0)
		return (0);
	else
	{
		lTemp = 255 - ((255 - Color2) * (255 - Color2)) / Color1;
		return ((lTemp < 0) ? 0 : (BYTE)lTemp);
	}
}

// Additive blend mode
inline BYTE AdditiveMode (BYTE Color1, BYTE Color2)
{
	int lTemp;

	lTemp = Color1 + Color2;
	return ((lTemp > 255) ? 255 : (BYTE)lTemp);
}

// Subtractive blend mode
inline BYTE SubtractiveMode (BYTE Color1, BYTE Color2)
{
	int lTemp;

	lTemp = Color1 + Color2 - 256;
	return ((lTemp < 0) ? 0 : (BYTE)lTemp);
}

// Interpolation blend mode
inline BYTE InterpolationMode (BYTE Color1, BYTE Color2)
{
	int lTemp;

	Color1 = (BYTE)(64 - cos (Color1 * 3.14 / 255) * 64 + 0.5);
	Color2 = (BYTE)(64 - cos (Color2 * 3.14 / 255) * 64 + 0.5);
	
	lTemp = Color1 + Color2;
	return ((lTemp > 255) ? 255 : (BYTE)lTemp);
}

// Stamp blend mode
inline BYTE StampMode (BYTE Color1, BYTE Color2)
{
	int lTemp;

	lTemp = Color1 + 2 * Color2 - 256;
	return ((lTemp < 0) ? 0 : (lTemp > 255) ? 255 : (BYTE)lTemp);
}

// XOR blend mode
inline BYTE XORMode (BYTE Color1, BYTE Color2)
{
	return (Color1 ^ Color2);
}

// This is the function that call all the other blend modes.
// I use this function in the GPX_BlendMode function.
// Only you have to do is pass the two colors and the blend mode.
BYTE ApplyBlendMode (BYTE Color1, BYTE Color2, int Mode)
{
	switch (Mode)
	{
		case BLM_AVERAGE:
			return (AverageMode (Color1, Color2));
			break;
		case BLM_MULTIPLY:
			return (MultiplyMode (Color1, Color2));
			break;
		case BLM_SCREEN:
			return (ScreenMode (Color1, Color2));
			break;
		case BLM_DARKEN:
			return (DarkenMode (Color1, Color2));
			break;
		case BLM_LIGHTEN:
			return (LightenMode (Color1, Color2));
			break;
		case BLM_DIFFERENCE:
			return (DifferenceMode (Color1, Color2));
			break;
		case BLM_NEGATION:
			return (NegationMode (Color1, Color2));
			break;
		case BLM_EXCLUSION:
			return (ExclusionMode (Color1, Color2));
			break;
		case BLM_OVERLAY:
			return (OverlayMode (Color1, Color2));
			break;
		case BLM_HARDLIGHT:
			return (HardLightMode (Color1, Color2));
			break;
		case BLM_SOFTLIGHT:
			return (SoftLightMode (Color1, Color2));
			break;
		case BLM_COLORDODGE:
			return (ColorDodgeMode (Color1, Color2));
			break;
		case BLM_COLORBURN:
			return (ColorBurnMode (Color1, Color2));
			break;
		case BLM_SOFTDODGE:
			return (SoftDodgeMode (Color1, Color2));
			break;
		case BLM_SOFTBURN:
			return (SoftBurnMode (Color1, Color2));
			break;
		case BLM_REFLECT:
			return (ReflectMode (Color1, Color2));
			break;
		case BLM_GLOW:
			return (GlowMode (Color1, Color2));
			break;
		case BLM_FREEZE:
			return (FreezeMode (Color1, Color2));
			break;
		case BLM_HEAT:
			return (HeatMode (Color1, Color2));
			break;
		case BLM_ADDITIVE:
			return (AdditiveMode (Color1, Color2));
			break;
		case BLM_SUBTRACTIVE:
			return (SubtractiveMode (Color1, Color2));
			break;
		case BLM_INTERPOLATION:
			return (InterpolationMode (Color1, Color2));
			break;
		case BLM_STAMP:
			return (StampMode (Color1, Color2));
			break;
		case BLM_XOR:
			return (XORMode (Color1, Color2));
			break;
		default:
			return (Color1);
			break;
	}
}