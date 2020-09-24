/************************************************************************************
 ********                                                                    ********
 ********                                INTRO                               ********
 ********                                                                    ********
 ************************************************************************************/

// GraphicalDLL by Pieter Z. Voloshyn
//
// This is the GraphicalDLL source code, I try to explain what my functions are doing
// but is very difficult to abstract all the functions, ok?
//
// If you want to use this code in your application, don't forget to mention that's
// my code, or that's based on my code, all right?
//
// Ok! Here we go...
// This is a type library that simplify the C++ -> Visual Basic communication (see the
// ODL code). You can use a module with all this functions, but is a bit more difficult.
//
// To use in VB, simply compile this code (I use VC++ 6.0). If you compile in Release
// built, this code has about 144 Kb.
// After this, place the dll in your VB project directory.
//
// In VB enviroment, simply go to Project Menu, References... and browse to the 
// GraphicalDLL.dll, a piece of cake, hun?
// You can use the VB module with all the declarations instead reference the DLL
//
// If you press F2 in VB enviroment, the Object Browser will appear and you can see
// all the functions, parameters, etc...
//
// Note that you don't need to use the Response parameter, VB will do all the dirty
// work for you. The return (HRESULT) will be used by the system, not your application.
//
// If you want to use a module, you have to pass the Response parameter (byref), ok?
//

/************************************************************************************
 ********                                                                    ********
 ********                               HEADERS                              ********
 ********                                                                    ********
 ************************************************************************************/

// Headers to be used in GraphicalDLL
#include "stdafx.h"
#include "GraphicalDLL.h"

/************************************************************************************
 ********                                                                    ********
 ********                           DLL ENTRY-POINT                          ********
 ********                                                                    ********
 ************************************************************************************/
//HMODULE hInstance;

/* This is the entry-point for the DLL												*/
BOOL WINAPI DllMain (HANDLE	hModule, 
					 DWORD	ul_reason_for_call, 
					 LPVOID	lpReserved)
{
	return (TRUE);
}



/************************************************************************************
 ********                                                                    ********
 ********                        EXPORTED FUNCTIONS                          ********
 ********                                                                    ********
 ************************************************************************************/

/* Function to replace a colour by another 											*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * SelectColor		=> Colour to be removed											*
 * SubstituteColor  => Colour to substitute											*
 * Range			=> Works like a sensibility										*
 *																					*
 * Theory			=> This function replace a well defined color by another, but	*
 *					changing the Range, will increase or decrease the colors range	*
 *                  E.g. with a color 100, 50, 30 and range 20, we have a variation *
 *					between 75, 25, 5 to 125, 75, 55. Its a lot of colors, hun?		*
 *																					*/
HRESULT __stdcall GPX_BackDropRemoval (HDC		PicDestDC, 
									   HDC		PicSrcDC, 
									   UINT		SelectColor, 
									   UINT		SubstituteColor,
									   int		Range,
									   int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	int TempVar		= 0;
	BITMAPINFO info;
	// Here, we get the bitmap handle from our hdc
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	// Here we extracts from a integer color the R, G and B values
	BYTE SelectR = (SelectColor & 0x000000FF),
		 SelectG = (SelectColor & 0x0000FF00) >> 8,
		 SelectB = (SelectColor & 0x00FF0000) >> 16;

	BYTE SubsR = (SubstituteColor & 0x000000FF),
		 SubsG = (SubstituteColor & 0x0000FF00) >> 8,
		 SubsB = (SubstituteColor & 0x00FF0000) >> 16;

	// if bimap handle doesn't exists...
	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	// I will explain one time for this calc...
	// LineWidth is the pixel's width, each color has 3 components (R, G and B)
	// Stride is the offset between one scan line and the next scan line
 	int LineWidth = Width * 3;
	int Stride = 4 - LineWidth % 4;
	if (LineWidth % 4)
		LineWidth += Stride;

	// Total of bytes
	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		// This part calcule all the possible ranges (mins and maxs)
		TempVar = (int)(SelectR - ((Range / 2) * COLOR_PERCENTAGE));
		int RangeMinR = ((TempVar > 0) ? TempVar : 0);
		TempVar = (int)(SelectG - ((Range / 2) * COLOR_PERCENTAGE));
		int RangeMinG = ((TempVar > 0) ? TempVar : 0);
		TempVar = (int)(SelectB - ((Range / 2) * COLOR_PERCENTAGE));
		int RangeMinB = ((TempVar > 0) ? TempVar : 0);
		TempVar = (int)(SelectR + ((Range / 2) * COLOR_PERCENTAGE));
		int RangeMaxR = ((TempVar < 255) ? TempVar : 255);
		TempVar = (int)(SelectG + ((Range / 2) * COLOR_PERCENTAGE));
		int RangeMaxG = ((TempVar < 255) ? TempVar : 255);
		TempVar = (int)(SelectB + ((Range / 2) * COLOR_PERCENTAGE));
		int RangeMaxB = ((TempVar < 255) ? TempVar : 255);
			
		int i = 0;
		for (int h = 0; h < Height; h++, i += Stride)
			for (int w = 0; w < Width; w++, i += 3)
			{
				if (((Bits[i+2] >= RangeMinR) && (Bits[i+2] <= RangeMaxR)) &&
					((Bits[i+1] >= RangeMinG) && (Bits[i+1] <= RangeMaxG)) &&
					((Bits[ i ] >= RangeMinB) && (Bits[ i ] <= RangeMaxB)))
				{
					Bits[i+2] = SubsR;	// If the color is between the ranges
					Bits[i+1] = SubsG;	// replace by another
					Bits[ i ] = SubsB;
				}
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to remove the back of an image finding edges 							*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * SelectColor		=> Colour to be removed											*
 * SubstituteColor  => Colour to substitute											*
 * Range			=> Works like a sensibility										*
 * Top				=> Scan from top to botton										*
 * Left				=> Scan from left to right										*
 * Right			=> Scan from right to left										*
 * Botton			=> Scan from botton to top										*
 *																					*
 * Theory			=> Do you ever seen LeadTools functions? Here I create a more	*
 *					fast code to analize the borders.								*
 *					Imagine a photo 3X4 (used in documents, etc...) and the			*
 *					back of the photo is blue, but you want to make this back white	*
 *					but you think - I'll use the replace color to make this -, but	*
 *					you're a very unlucky guy and the photo employee have a blue	*
 *					tie, this will be replaced by another color.					*
 *					This function will analize for borders.							*
 *																					*/
HRESULT __stdcall GPX_BackDropRemovalEx (HDC	PicDestDC, 
										 HDC	PicSrcDC, 
										 UINT	SelectColor, 
										 UINT	SubstituteColor,
										 int	Range, 
										 BOOL	Top, 
										 BOOL	Left, 
										 BOOL	Right, 
										 BOOL	Botton,
										 int	*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	int TempVar		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	BYTE SelectR = (SelectColor & 0x000000FF),
		 SelectG = (SelectColor & 0x0000FF00) >> 8,
		 SelectB = (SelectColor & 0x00FF0000) >> 16,

		 SubsR = (SubstituteColor & 0x000000FF),
		 SubsG = (SubstituteColor & 0x0000FF00) >> 8,
		 SubsB = (SubstituteColor & 0x00FF0000) >> 16;

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	int LineWidth = Width * 3;
	int Stride = 4 - LineWidth % 4;
	if (LineWidth % 4)
		LineWidth += Stride;

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];
	LPBYTE Flags = new BYTE[BitCount];
	ZeroMemory (Flags, BitCount);

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		// This part calcule the ranges (mins and maxs)
		TempVar = (int)(SelectR - ((Range / 2) * COLOR_PERCENTAGE));
		int RangeMinR = ((TempVar > 0) ? TempVar : 0);
		TempVar = (int)(SelectG - ((Range / 2) * COLOR_PERCENTAGE));
		int RangeMinG = ((TempVar > 0) ? TempVar : 0);
		TempVar = (int)(SelectB - ((Range / 2) * COLOR_PERCENTAGE));
		int RangeMinB = ((TempVar > 0) ? TempVar : 0);
		TempVar = (int)(SelectR + ((Range / 2) * COLOR_PERCENTAGE));
		int RangeMaxR = ((TempVar < 255) ? TempVar : 255);
		TempVar = (int)(SelectG + ((Range / 2) * COLOR_PERCENTAGE));
		int RangeMaxG = ((TempVar < 255) ? TempVar : 255);
		TempVar = (int)(SelectB + ((Range / 2) * COLOR_PERCENTAGE));
		int RangeMaxB = ((TempVar < 255) ? TempVar : 255);

		int i, h, w;

		// Here the top->botton scan
		if (Top)
		{
			i = 0;
			for (w = 0; w < Width; w++)
				for (h = 0; h < Height; h++)
				{
					i = (Height - h - 1) * LineWidth + 3 * w;
					if ((Bits[i+2] == SubsR) && (Bits[i+1] == SubsG) && (Bits[i] == SubsB))
						continue;
					else
					{
						if (((Bits[i+2] >= RangeMinR) && (Bits[i+2] <= RangeMaxR)) &&
						    ((Bits[i+1] >= RangeMinG) && (Bits[i+1] <= RangeMaxG)) &&
						    ((Bits[ i ] >= RangeMinB) && (Bits[ i ] <= RangeMaxB)))
							Flags[i] = 1;
						else
							h = Height;
					}
				}
		}

		// Here the left->right scan
		if (Left)
		{
			i = 0;
			for (h = 0; h < Height; h++)
				for (w = 0; w < Width; w++)
				{
					i = h * LineWidth + 3 * w;
					if ((Bits[i+2] == SubsR) && (Bits[i+1] == SubsG) && (Bits[i] == SubsB))
						continue;
					else
					{
						if (((Bits[i+2] >= RangeMinR) && (Bits[i+2] <= RangeMaxR)) &&
						    ((Bits[i+1] >= RangeMinG) && (Bits[i+1] <= RangeMaxG)) &&
						    ((Bits[ i ] >= RangeMinB) && (Bits[ i ] <= RangeMaxB)))
							Flags[i] = 1;
						else
							w = Width;
					}
				}
		}

		// Here the right->left scan
		if (Right)
		{
			i = 0;
			for (h = 0; h < Height; h++)
				for (w = 0; w < Width; w++)
				{
					i = h * LineWidth + 3 * (Width - w - 1);
					if ((Bits[i+2] == SubsR) && (Bits[i+1] == SubsG) && (Bits[i] == SubsB))
						continue;
					else
					{
						if (((Bits[i+2] >= RangeMinR) && (Bits[i+2] <= RangeMaxR)) &&
						    ((Bits[i+1] >= RangeMinG) && (Bits[i+1] <= RangeMaxG)) &&
						    ((Bits[ i ] >= RangeMinB) && (Bits[ i ] <= RangeMaxB)))
							Flags[i] = 1;
						else
							w = Width;
					}
				}
		}

		// Here the botton->top scan
		if (Botton)
		{
			i = 0;
			for (w = 0; w < Width; w++)
				for (h = 0; h < Height; h++)
				{
					i = h * LineWidth + 3 * w;
					if ((Bits[i+2] == SubsR) && (Bits[i+1] == SubsG) && (Bits[i] == SubsB))
						continue;
					else
					{
						if ((Bits[i+2] >= RangeMinR) && (Bits[i+2] <= RangeMaxR) &&
						    (Bits[i+1] >= RangeMinG) && (Bits[i+1] <= RangeMaxG) &&
						    (Bits[ i ] >= RangeMinB) && (Bits[ i ] <= RangeMaxB))
							Flags[i] = 1;
						else
							h = Height;
					}
				}
		}

		i = 0;
		for (h = 0; h < Height; h++, i += Stride)
			for (w = 0; w < Width; w++, i += 3)
			{
				if (Flags[i] == 1)
				{
					Bits[i+2] = SubsR;
					Bits[i+1] = SubsG;
					Bits[i] = SubsB;
				}
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;
		delete [] Flags;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the Sepia effect 												*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 *																					*
 * Theory			=> This is a great effect, in PSC I don't seen this effect and	*
 *					its very simple. Firstly, we change the color to the grayscale	*
 *					After this, we make a "brownscale" huahua using simple adds		*
 *					and subs (+ and -)												*
 *																					*/
HRESULT __stdcall GPX_Sepia (HDC		PicDestDC, 
							 HDC		PicSrcDC,
							 int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0;
		int GrayPixel;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				// Get the gray tone
				GrayPixel = ((Bits[i+2] + Bits[i+1] + Bits[i]) / 3);
				// Make a sepia tone
				Bits[i+2] = (BYTE)(GrayPixel > 202) ? 255 : GrayPixel + 53;
				Bits[i+1] = (BYTE)(GrayPixel > 235) ? 255 : GrayPixel + 20;
				Bits[ i ] = (BYTE)(GrayPixel <  33) ?   0 : GrayPixel - 33;
				Bits[i+1] = (Bits[i+1] < 30) ? 30 : Bits[i+1];
				Bits[ i ] = (Bits[ i ] < 30) ? 30 : Bits[ i ];
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to reduce the image to only 2 colours 									*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 *																					*
 * Theory			=> This is a very simple effect. We analize the color, if less	*
 *					than 128 ? Yes ? So, this color will be black... No ? So, this	*
 *					color will be white... easy, hun?								*
 *																					*/
HRESULT __stdcall GPX_ReduceTo2Colors (HDC		PicDestDC, 
									   HDC		PicSrcDC,
									   int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0;
		int GrayPixel;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				// Get the gray tone
				GrayPixel = ((Bits[i+2] + Bits[i+1] + Bits[i]) / 3);
				GrayPixel = (GrayPixel > 127) ? 255 : 0;
				Bits[i+2] = Bits[i+1] = Bits[ i ] = (BYTE)GrayPixel;
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to reduce to 8 colours													*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 *																					*
 * Theory			=> This function is very similar to GPX_ReduceTo2Colors, but	*
 *					here we don't need to generate the gray tone. We analize each	*
 *					8 bits. E.g. R = 198, G = 26, B = 38 with this function, this	*
 *					pixel will be R = 255, G = 0, B = 0.							*
 *																					*/
HRESULT __stdcall GPX_ReduceTo8Colors (HDC		PicDestDC, 
									   HDC		PicSrcDC,
									   int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		UCHAR c_Table[COLOR_SIZE];
		for (int i = 0; i < 256; i++)
			c_Table[i] = (i > 127) ? 255 : 0;
		
		AssignTables (c_Table, c_Table, c_Table, Bits, Width, Height);

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to reduce the image colours												*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Levels			=> Value to set the level of colour	reduction					*
 *																					*
 * Theory			=> Here, the level indicates the reduction. We calcule with		*
 *					this formula: color = color - (color % level).					*
 *					E.g. R = 187 and level = 5, so, R will be 187 - (187 % 5) = 185	*
 *																					*/
HRESULT __stdcall GPX_ReduceColors (HDC		PicDestDC, 
									HDC		PicSrcDC,
									int		Levels,
									int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (Levels < 0)
		Levels = 0;

	Levels++;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		UCHAR rc[COLOR_SIZE];
		for (int i = 0; i < 256; i++)
			rc[i] = LimitValues (i - (i % Levels));

		AssignTables (rc, rc, rc, Bits, Width, Height);

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the Stamp effect 												*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Level			=> Sensibility value											*
 *																					*
 * Theory			=> Here, we have a function with the same logic as to reduce	*
 *					to 2 colors, but now, I made a sensibility instead 128.			*
 *																					*/
HRESULT __stdcall GPX_Stamp (HDC		PicDestDC, 
							 HDC		PicSrcDC,
							 int		Level,
							 int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (Level < 0)
		Level = 0;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0;
		int GrayPixel;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				GrayPixel = ((Bits[i+2] + Bits[i+1] + Bits[i]) / 3);
				GrayPixel = (GrayPixel > Level) ? 255 : 0;
				Bits[i+2] = Bits[i+1] = Bits[ i ] = (BYTE)GrayPixel;
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to adjust the image brightness 											*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Value			=> Brightness value												*
 *																					*
 * Theory			=> This is one of the most easy functions to understand, this	*
 *					functions takes a pixel and add or sub value to that pixel.		*
 *					As you increase the value, you make the image more bright		*
 *																					*/
HRESULT __stdcall GPX_Brightness (HDC		PicDestDC, 
								  HDC		PicSrcDC,
								  int		Value,
								  int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = ::GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		UCHAR brightTable[COLOR_SIZE];
		for (int i = 0; i < 256; i++)
			brightTable[i] = LimitValues (i + Value);

		AssignTables (brightTable, brightTable, brightTable, Bits, Width, Height);

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the Rock effect 												*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Value			=> Intensity of rock effect										*
 *																					*
 * Theory			=> I made this effect accidently. I was doing the sharpen		*
 *					effect, but I only inverse the order in one line to make this	*
 *					effect. This function, takes a pixel an the next diagonal pixel	*
 *					and calcule the intensity (like a border detect)				* 
 *																					*/
HRESULT __stdcall GPX_Rock (HDC		PicDestDC, 
							HDC		PicSrcDC,
							int		Value,
							int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (Value < 0)
		Value = 0;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0;
		for (int h = 1; h < Height; h++)
			for (int w = 1; w < Width; w++)
			{
				j = h * LineWidth + 3 * w;
				i = (h - 1) * LineWidth + 3 * (w - 1);
				Bits[i+2] = LimitValues ((int)Bits[i+2] + (Value * (Bits[i+2] - Bits[j+2])));
				Bits[i+1] = LimitValues ((int)Bits[i+1] + (Value * (Bits[i+1] - Bits[j+1])));
				Bits[ i ] = LimitValues ((int)Bits[ i ] + (Value * (Bits[ i ] - Bits[ j ])));
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to adjust the image sharpen												*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Value			=> Sharpen value												*
 *																					*
 * Theory			=> Similar to Rock effect, but has the calcule inversed.		*
 *					With this theory, you can do another sharpen effects like		*
 *					SharpenMore, SharpBorders and so on...							*
 *																					*/
HRESULT __stdcall GPX_Sharpening (HDC		PicDestDC, 
								  HDC		PicSrcDC,
								  float		Value,
								  int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (Value < 0)
		Value = 0;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				j = abs((h - 1) * LineWidth + 3 * (w - 1));
				Bits[i+2] = LimitValues ((int)Bits[i+2] + ((int)(Value * (Bits[i+2] - Bits[j+2]))));
				Bits[i+1] = LimitValues ((int)Bits[i+1] + ((int)(Value * (Bits[i+1] - Bits[j+1]))));
				Bits[ i ] = LimitValues ((int)Bits[ i ] + ((int)(Value * (Bits[ i ] - Bits[ j ]))));
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the AmbientLight effect (based on Martin code) 				*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * AmbientColor		=> Ambient color												*
 * Intensity		=> Light Intensity												*
 *																					*
 * Theory			=> This is a great effect, there is a calcule that defines the	*
 *					color luminance. Best viewed with a medium to dark tone			* 
 *																					*/
HRESULT __stdcall GPX_AmbientLight (HDC		PicDestDC, 
									HDC		PicSrcDC,
									int		AmbientColor, 
									int		Intensity,
									int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	Intensity = LimitValues (Intensity);
	
	BYTE LightR = LimitValues ((int)255 - Intensity -  (AmbientColor & 0x000000FF)),
		 LightG = LimitValues ((int)255 - Intensity - ((AmbientColor & 0x0000FF00) >>  8)),
		 LightB = LimitValues ((int)255 - Intensity - ((AmbientColor & 0x00FF0000) >> 16));

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				Bits[i+2] = LimitValues ((int)Bits[i+2] - LightR);
				Bits[i+1] = LimitValues ((int)Bits[i+1] - LightG);
				Bits[ i ] = LimitValues ((int)Bits[ i ] - LightB);
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the antialias effect 											*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 *																					*
 * Theory			=> This is a blur with less pixels.								*
 *																					*/
HRESULT __stdcall GPX_AntiAlias (HDC	PicDestDC, 
								 HDC	PicSrcDC,
								 int	Sensibility,
								 int	*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	Sensibility = LimitValues (Sensibility);

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];
	BYTE Temp[3][9];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0, k = 0, GrayCmp, Gray, add;
		for (int h = 1; h < Height - 1; h++)
			for (int w = 1; w < Width - 1; w++)
			{
				i = h * LineWidth + 3 * w;
				j = (h + 1) * LineWidth + 3 * w;
				k = (h - 1) * LineWidth + 3 * w;
				Gray = (Bits[i+2] + Bits[i+1] + Bits[i]) / 3;

				for (int y = 0; y < 3; y++)
				{
					add = y * 3;
					GrayCmp = (Bits[j+add-1] + Bits[j+add-2] + Bits[j+add-3]) / 3;
					if (! ((GrayCmp > Gray + Sensibility) || (GrayCmp < Gray - Sensibility)))
					{
						Temp[0][y] = Bits[ i ];
						Temp[1][y] = Bits[i+1];
						Temp[2][y] = Bits[i+2];
					}
					else
					{
						Temp[0][y] = Bits[j+add-3];
						Temp[1][y] = Bits[j+add-2];
						Temp[2][y] = Bits[j+add-1];
					}
				}

				for (y = 0; y < 3; y++)
				{
					add = y * 3;
					GrayCmp = (Bits[i+add-1] + Bits[i+add-2] + Bits[i+add-3]) / 3;
					if (! ((GrayCmp > Gray + Sensibility) || (GrayCmp < Gray - Sensibility)))
					{
						Temp[0][y+3] = Bits[ i ];
						Temp[1][y+3] = Bits[i+1];
						Temp[2][y+3] = Bits[i+2];
					}
					else
					{
						Temp[0][y+3] = Bits[i+add-3];
						Temp[1][y+3] = Bits[i+add-2];
						Temp[2][y+3] = Bits[i+add-1];
					}
				}

				for (y = 0; y < 3; y++)
				{
					add = y * 3;
					GrayCmp = (Bits[k+add-1] + Bits[k+add-2] + Bits[k+add-3]) / 3;
					if (! ((GrayCmp > Gray + Sensibility) || (GrayCmp < Gray - Sensibility)))
					{
						Temp[0][y+6] = Bits[ i ];
						Temp[1][y+6] = Bits[i+1];
						Temp[2][y+6] = Bits[i+2];
					}
					else
					{
						Temp[0][y+6] = Bits[k+add-3];
						Temp[1][y+6] = Bits[k+add-2];
						Temp[2][y+6] = Bits[k+add-1];
					}
				}

				Bits[i+2] = (Temp[2][0] + Temp[2][1] + Temp[2][2] +
							 Temp[2][3] + Temp[2][4] + Temp[2][5] +
							 Temp[2][6] + Temp[2][7] + Temp[2][8]) / 9;
				Bits[i+1] = (Temp[1][0] + Temp[1][1] + Temp[1][2] +
							 Temp[1][3] + Temp[1][4] + Temp[1][5] +
							 Temp[1][6] + Temp[1][7] + Temp[1][8]) / 9;
				Bits[ i ] = (Temp[0][0] + Temp[0][1] + Temp[0][2] +
							 Temp[0][3] + Temp[0][4] + Temp[0][5] +
							 Temp[0][6] + Temp[0][7] + Temp[0][8]) / 9;
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the Emboss effect 												*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Depth			=> Emboss value													*
 *																					*
 * Theory			=> This is an amazing effect. And the theory is very simple to	*
 *					understand. You get the diference between the colors and		*
 *					increase it. After this, get the gray tone						*
 *																					*/
HRESULT __stdcall GPX_Emboss (HDC	  PicDestDC, 
							  HDC	  PicSrcDC, 
							  float	  Depth,
							  int	  *Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0;
		int R = 0, G = 0, B = 0;
		BYTE Gray = 0;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				j = (h + Lim_Max (h, 1, Height)) * LineWidth + 3 * (w + Lim_Max (w, 1, Width));

				R = abs ((int)((Bits[i+2] - Bits[j+2]) * Depth + 128));
				G = abs ((int)((Bits[i+1] - Bits[j+1]) * Depth + 128));
				B = abs ((int)((Bits[ i ] - Bits[ j ]) * Depth + 128));

				Gray = LimitValues ((R + G + B) / 3);

				Bits[i+2] = Gray;
				Bits[i+1] = Gray;
				Bits[ i ] = Gray;
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the gamma effect 												*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Gamma			=> Gamma value													*
 *																					*
 * Theory			=> Applying this function, the effect is similar to brightness	*
 *					but light tones don't change too much							*
 *																					*/
HRESULT __stdcall GPX_Gamma (HDC		PicDestDC, 
							 HDC		PicSrcDC, 
							 float		Gamma,
							 int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (Gamma < 0)
		Gamma = 0;
	if (Gamma > 10)
		Gamma = 10;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		UCHAR *gTable = new UCHAR[COLOR_SIZE];
		for (int i = 0; i < 256; i++)
			gTable[i] = (BYTE)(255.0 * pow((i / 255.0), (1.0 / Gamma)));

		AssignTables (gTable, gTable, gTable, Bits, Width, Height);

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;
		delete [] gTable;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the Invert effect 												*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Intensity		=> Value for invertion											*
 *																					*
 * Theory			=> To invert a color is very easy, just take this formula in	*
 *					your mind: 255 (max value for a RGB value) - color.				*
 *					but here, we can make this more interesting with Intensity value*
 *																					*/
HRESULT __stdcall GPX_Invert (HDC		PicDestDC, 
							  HDC		PicSrcDC, 
							  int		Intensity,
							  int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (Intensity > 255)
		Intensity = 255;
	if (Intensity < 0)
		Intensity = 0;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;

				Bits[i+2] = LimitValues (abs((int)Intensity - Bits[i+2]));
				Bits[i+1] = LimitValues (abs((int)Intensity - Bits[i+1]));
				Bits[ i ] = LimitValues (abs((int)Intensity - Bits[ i ]));
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the Shift effect 												*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Shift			=> Shift value													*
 *																					*
 * Theory			=> This is a interesting effect, and easy to undestand. The		*
 *					process make a softly swap with blue and green tones.			*
 *																					*/
HRESULT __stdcall GPX_Shift (HDC		PicDestDC, 
							 HDC		PicSrcDC,
							 int		Shift,
							 int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (Shift > 255)
		Shift = 255;
	if (Shift < 0)
		Shift = 0;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;

				Bits[ i ] = (BYTE)ShadeColors ((int)Bits[ i ], (int)Bits[i+1], Shift);
				Bits[i+1] = (BYTE)ShadeColors ((int)Bits[i+1], (int)Bits[ i ], Shift);
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the Tone effect (based on Martin code) 						*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Tone				=> Tone value													*
 *																					*
 * Theory			=> This is an easy effect too, the image colors will be			*
 *					reaching the defined color adjusting the tone parameter			*
 *																					*/
HRESULT __stdcall GPX_Tone (HDC		PicDestDC, 
							HDC		PicSrcDC,
							int		Color, 
							int		Tone,
							int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (Tone > 255)
		Tone = 255;
	if (Tone < 0)
		Tone = 0;

	Tone = LimitValues (255 - Tone);

	BYTE R = (Color & 0x000000FF),
				  G = (Color & 0x0000FF00) >>  8,
				  B = (Color & 0x00FF0000) >> 16;

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;

				Bits[i+2] = (BYTE)ShadeColors ((int)R, (int)Bits[i+2], Tone);
				Bits[i+1] = (BYTE)ShadeColors ((int)G, (int)Bits[i+1], Tone);
				Bits[ i ] = (BYTE)ShadeColors ((int)B, (int)Bits[ i ], Tone);
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to adjust the Contrast 													*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * R				=> Red contrast adjustment										*
 * G				=> Green contrast adjustment									*
 * B				=> Blue contrast adjustment										*
 *																					*
 * Theory			=> Before the image processing, is created a table contents		*
 *					all the colors with contrast made. In the image processing		*
 *					the color is changed with the table color						*
 *																					*/
HRESULT __stdcall GPX_Contrast (HDC		PicDestDC, 
								HDC		PicSrcDC,
								float	Red, 
								float	Green, 
								float	Blue,
								int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (Red   < 0)   Red = 0;
	if (Green < 0) Green = 0;
	if (Blue  < 0)  Blue = 0;

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0;
		BYTE TableR[COLOR_SIZE], TableG[COLOR_SIZE], TableB[COLOR_SIZE];
		for (i; i < 256; i++)
		{
			TableR[i] = LimitValues ((int)((i - 127) *   Red) + 127);
			TableG[i] = LimitValues ((int)((i - 127) * Green) + 127);
			TableB[i] = LimitValues ((int)((i - 127) *  Blue) + 127);
		}

		AssignTables (TableR, TableG, TableB, Bits, Width, Height);
		
		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the GrayScale effect 											*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Scale			=> Bright value for the gray scale								*
 *																					*
 * Theory			=> To find the gray tone is very simple, apply this formula:	*
 *					(R+G+B)/3, easy hun?											*
 *																					*/
HRESULT __stdcall GPX_GrayScale (HDC		PicDestDC, 
								 HDC		PicSrcDC,
								 int		Scale,
								 int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (Scale < -255)
		Scale = -255;
	if (Scale > 255)
		Scale = 255;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0;
		int GrayPixel;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				GrayPixel = ((Bits[i+2] + Bits[i+1] + Bits[i]) / 3);
				if (Scale >= 0)
					GrayPixel = LimitValues (GrayPixel + Scale);
				else
					GrayPixel = ShadeColors (GrayPixel, 0x00000000, abs (Scale));
				Bits[i+2] = Bits[i+1] = Bits[ i ] = (BYTE)GrayPixel;
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the RandomicalPoints effect 									*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * RandValue		=> Max value for random value									*
 * BackColor		=> Color used for the random points								*
 *																					*
 * Theory			=> This effect is very simple to understand, the points will	*
 *					appear in a random processing									*
 *																					*/
HRESULT __stdcall GPX_RandomicalPoints (HDC		PicDestDC, 
										HDC		PicSrcDC,
										int		RandValue, 
										int		BackColor,
										int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	BYTE BackR = (BackColor & 0x000000FF),
				  BackG = (BackColor & 0x0000FF00) >> 8,
				  BackB = (BackColor & 0x00FF0000) >> 16;

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (RandValue < 1)
		RandValue = 1;
	if (RandValue > 255)
		RandValue = 255;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		srand ((UINT) GetTickCount());
		int i = 0, r;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				r = (rand() % RandValue) + 1;
				if ((r == 2) || (r == 1))
				{
					Bits[i+2] = BackR;
					Bits[i+1] = BackG;
					Bits[ i ] = BackB;
				}
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the ColorRandomize effect 										*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * RandValue		=> Value to be added in each pixel								*
 *																					*
 * Theory			=> The RandValue will be added in a random choice: R, G or B	*
 *																					*/
HRESULT __stdcall GPX_ColorRandomize (HDC		PicDestDC, 
									  HDC		PicSrcDC,
									  int		RandValue,
									  int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (RandValue < -255)
		RandValue = -255;
	if (RandValue > 255)
		RandValue = 255;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		srand ((UINT) GetTickCount());
		int i = 0, r;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				r = (rand() % 3) + 1;
				if (r == 1)
					Bits[i+2] = LimitValues (Bits[i+2] + RandValue);
				else if (r == 2)
					Bits[i+1] = LimitValues (Bits[i+1] + RandValue);
				else
					Bits[ i ] = LimitValues (Bits[ i ] + RandValue);
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the Solarize effect 											*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Invert			=> If inverts the result										*
 *																					*
 * Theory			=> Similar to invert effect, but here, only dark tones will		*
 *					be inverted														*
 *																					*/
HRESULT __stdcall GPX_Solarize (HDC		PicDestDC, 
								HDC		PicSrcDC,
								BOOL	Invert,
								int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				if (! Invert)
				{
					if (Bits[i+2] < 128)
						Bits[i+2] = 255 - Bits[i+2];
					if (Bits[i+1] < 128)
						Bits[i+1] = 255 - Bits[i+1];
					if (Bits[ i ] < 128)
						Bits[ i ] = 255 - Bits[ i ];
				}
				else
				{
					if (Bits[i+2] > 127)
						Bits[i+2] = 255 - Bits[i+2];
					if (Bits[i+1] > 127)
						Bits[i+1] = 255 - Bits[i+1];
					if (Bits[ i ] > 127)
						Bits[ i ] = 255 - Bits[ i ];
				}
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the Diffuse effect												*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 *																					*
 * Theory			=> The pixel will be moved to another random pixel, easy?		*
 *																					*/
HRESULT __stdcall GPX_Diffuse (HDC	PicDestDC, 
							   HDC	PicSrcDC,
							   int	*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0, rx, ry;
		srand ((UINT) GetTickCount());
		for (int h = 2; h < Height - 3; h++)
			for (int w = 2; w < Width - 3; w++)
			{
				rx = (rand() % 4) - 2;
				ry = (rand() % 4) - 2;
				i = h * LineWidth + 3 * w;
				j = (h + ry) * LineWidth + 3 * (w + rx);
				Bits[i+2] = Bits[j+2];
				Bits[i+1] = Bits[j+1];
				Bits[ i ] = Bits[ j ];
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the Mosaic effect 												*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Size				=> Size of mosaic (duhhh!!!)									*
 *																					*
 * Theory			=> Ok, you can find some mosaic effects on PSC, but this one	*
 *					has a great feature, if you see a mosaic in other code you will	*
 *					see that the corner pixel doesn't change. The explanation is	*
 *					simple, the color of the mosaic is the same as the first pixel	*
 *					get. Here, the color of the mosaic is the same as the mosaic	*
 *					center pixel													*
 *																					*/
HRESULT __stdcall GPX_Mosaic (HDC	PicDestDC, 
							  HDC	PicSrcDC,
							  int	Size,
							  int	*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (Size <= 0)
	{
		*Response = (int)Result;
		return (S_OK);
	}

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0, H, W;
		for (int h = 0; h < Height; h++)
		{
			for (int w = 0; w < Width; w++)
			{
				i = (h + (Lim_Max (h, Size, Height) / 2))  * LineWidth + 3 * (w + (Lim_Max (w, Size, Width) / 2));
				for (H = 0; H < Size; H++)
					for (W = 0; W < Size; W++)
					{
						if (h + H >= Height)
						{
							if (w + W >= Width)
								j = h * LineWidth + 3 * w;
							else
								j = h * LineWidth + 3 * (w + W);
						}
						else
						{
							if (w + W >= Width)
								j = (h + H) * LineWidth + 3 * w;
							else
								j = (h + H) * LineWidth + 3 * (w + W);
						}
						Bits[j+2] = Bits[i+2];
						Bits[j+1] = Bits[i+1];
						Bits[ j ] = Bits[ i ];
					}
				w += W - 1;
			}
		h += H - 1;
		}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the Melt effect 												*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 *																					*
 * Theory			=> This is a silly effect, because in the image processing, the	*
 *					line is swaped with the superior line							*
 *																					*/
HRESULT __stdcall GPX_Melt (HDC	PicDestDC, 
							HDC	PicSrcDC,
							int	*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0;
		for (int h = 0; h < Height - 1; h += 2)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				j = (h + 1) * LineWidth + 3 * w;
				Bits[i+2] = Bits[j+2] - Bits[i+2] + (Bits[j+2] = Bits[i+2]);
				Bits[i+1] = Bits[j+1] - Bits[i+1] + (Bits[j+1] = Bits[i+1]);
				Bits[ i ] = Bits[ j ] - Bits[ i ] + (Bits[ j ] = Bits[ i ]);
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the FishEye effect												*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 *																					*
 * Theory			=> Huahuahua, this is a great effect if you take employee photos*
 *					Its pure trigonometry. I think if you study hard the code you	*
 *					understand very well. ;)										*
 *																					*/
HRESULT __stdcall GPX_FishEye (HDC	PicDestDC, 
							   HDC	PicSrcDC,
							   int	*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];
	LPBYTE NewBits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0, nw, nh;
		double Angle, Radius, rNew;
		double Radmax = sqrt (Width * Width + Height * Height) / 2;
		for (int w = 0; w < Width; w++)
			for (int h = 0; h < Height; h++)
			{
				nw = Width / 2 - w;
				nh = Height / 2 - h;
				Radius = sqrt (nw * nw + nh * nh);
				if (Radius < Radmax)
				{
					Angle = atan2 (nh, nw);
					rNew = Radius * Radius / Radmax;
					nw = (int)(Width / 2 + rNew * cos (Angle));
					nh = (int)(Height / 2 - rNew * sin (Angle));
					nw = (nw < 0) ? 0 : ((nw > Width) ? Width : nw);
					nh = (nh < 0) ? 0 : ((nh > Height) ? Height : nh);
					i = h * LineWidth + 3 * (Width - w - 1);
					j = nh * LineWidth + 3 * nw;
					NewBits[i+2] = Bits[j+2];
					NewBits[i+1] = Bits[j+1];
					NewBits[ i ] = Bits[ j ];
				}
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, NewBits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;
		delete [] NewBits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the Swirl effect 												*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Swirl			=> Swirl value													*
 *																					*
 * Theory			=> Like FishEye effect, it's very difficult to explain how can	*
 *					I reach on this effect, study hard for this. It's pure			*
 *					trigonometry. If you have spiral theorems, you will understand	*
 *					better, ok?														*
 *																					*/
HRESULT __stdcall GPX_Swirl (HDC PicDestDC, 
							 HDC PicSrcDC,
							 int Swirl,
							 int *Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	Swirl = (Swirl > 255) ? 255 : (Swirl < -255) ? -255 : Swirl;
	
	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];
	LPBYTE NewBits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0, nw, nh;
		double Angle, Radius, aNew;
		double Radmax = sqrt (Height * Height + Width * Width) / 2;
		for (int w = 0; w < Width; w++)
			for (int h = 0; h < Height; h++)
			{
				nw = Width / 2 - w;
				nh = Height / 2 - h;
				Radius = sqrt (nw * nw + nh * nh);
				if (Radius < Radmax)
				{
					Angle = atan2 (nh, nw);
					aNew = Angle + Radius / Swirl;
					nw = (int)(Width / 2 + Radius * cos (aNew));
					nh = (int)(Height / 2 - Radius * sin (aNew));
					nw = (nw < 0) ? 0 : ((nw >= Width) ? Width - 1 : nw);
					nh = (nh < 0) ? 0 : ((nh >= Height) ? Height - 1 : nh);
					i = h * LineWidth + 3 * w;
					j = nh * LineWidth + 3 * nw;
					NewBits[i+2] = Bits[j+2];
					NewBits[i+1] = Bits[j+1];
					NewBits[ i ] = Bits[ j ];
				}
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, NewBits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;
		delete [] NewBits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the Twirl effect 												*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Twirl			=> Twirl value													*
 *																					*
 * Theory			=> Take spiral studies, you will understand better, I'm studying*
 *					hard on this effect, because it's not too fast.					*
 *																					*/
HRESULT __stdcall GPX_Twirl (HDC PicDestDC, 
							 HDC PicSrcDC,
							 int Twirl,
							 int *Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	Twirl = (Twirl > 100) ? 100 : (Twirl < -100) ? -100 : Twirl;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];
	LPBYTE NewBits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		double half_w = (double)Width / 2.0,
			   half_h = (double)Height / 2.0;
		double twAngle = (double)Twirl / (half_w * 10.0);
		double Angle, NewAngle, AngleAcc, Radius, Radmax;
		double nw, nh;
		int i, j;

		//HRGN hEllipse = ::CreateEllipticRgn (0, 0, Width, Height);
		for (int w = 0; w < Width; w++)
			for (int h = 0; h < Height; h++)
			{
				i = h * LineWidth + 3 * w;
				nw = half_w - (double)w;
				nh = half_h - (double)h;

				Radius = sqrt (nw * nw + nh * nh);
				Angle = atan2 (nh, nw);
				Radmax = MaximumRadius (Height, Width, Angle);
				AngleAcc = twAngle * (-1.0 * (Radius - Radmax));

				//if (::PtInRegion (hEllipse, w, h))
				if (Radius < Radmax)
				{
					NewAngle = Angle + AngleAcc;
					nw = half_w - cos (NewAngle) * Radius;
					nh = half_h - sin (NewAngle) * Radius;
					nw = (nw < 0.0) ? 0.0 : ((nw >= Width) ? (double)(Width - 1) : nw);
					nh = (nh < 0.0) ? 0.0 : ((nh >= Height) ? (double)(Height - 1) : nh);
					j = (int)nh * LineWidth + 3 * (int)nw;
					NewBits[i+2] = Bits[j+2];
					NewBits[i+1] = Bits[j+1];
					NewBits[ i ] = Bits[ j ];
				}
				else
				{
					NewBits[i+2] = Bits[i+2];
					NewBits[i+1] = Bits[i+1];
					NewBits[ i ] = Bits[ i ];
				}
			}
		
		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, NewBits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		//::DeleteObject (hEllipse);
		delete [] Bits;
		delete [] NewBits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the Neon effect 												*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Intensity		=> Intensity value												*
 * BW				=> Border Width													*
 *																					*
 * Theory			=> Wow, this is a great effect, you've never seen a Neon effect	*
 *					like this on PSC. Is very similar to Growing Edges (photoshop)	*
 *					Some pictures will be very interesting							*
 *																					*/
HRESULT __stdcall GPX_Neon (HDC		PicDestDC, 
							HDC		PicSrcDC,
							short	Intensity, 
							short	BW,
							int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	Intensity = (Intensity < 0) ? 0 : (Intensity > 5) ? 5 : Intensity;
	BW = (BW < 1) ? 1 : (BW > 5) ? 5 : BW;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0, color_1, color_2;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
				for (int k = 0; k <= 2; k++)
				{
					i = h * LineWidth + 3 * w;
					j = h * LineWidth + 3 * (w + Lim_Max (w, BW, Width));
					color_1 = (int)((Bits[i+k] - Bits[j+k]) * (Bits[i+k] - Bits[j+k]));
					j = (h + Lim_Max (h, BW, Height)) * LineWidth + 3 * w;
					color_2 = (int)((Bits[i+k] - Bits[j+k]) * (Bits[i+k] - Bits[j+k]));
					Bits[i+k] = LimitValues ((int)(sqrt ((color_1 + color_2) << Intensity)));
				}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the Canvas effect 												*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Canvas			=> Canvas value													*
 *																					*
 * Theory			=> This is a nice effect, if you made the right change on Canvas*
 *					you'll see half-original, half-mirrored							*
 *																					*/
HRESULT __stdcall GPX_Canvas (HDC PicDestDC, 
							  HDC PicSrcDC,
							  int Canvas,
							  int *Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	if (Canvas > Width || Canvas == 0)
		return (-1);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];
	LPBYTE Stib = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0, h, w;
		for (h = 0; h < Height; h++)
			for (w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				j = h * LineWidth + 3 * (Width - w);
				Stib[i+2] = Bits[j+2];
				Stib[i+1] = Bits[j+1];
				Stib[ i ] = Bits[ j ];
			}
		
		for (w = Canvas; w < Width; w++)
			for (h = 0; h < Height; h++)
			{
				i = h * LineWidth + 3 * w;
				Bits[i+2] = Stib[i+2];
				Bits[i+1] = Stib[i+1];
				Bits[ i ] = Stib[ i ];
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;
		delete [] Stib;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the Waves effect 												*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Amplitude		=> Sinoidal maximum height										*
 * Frequency		=> Frequency value												*
 * FillSides		=> Like a boolean variable										*
 * Direction		=> Vertical or horizontal flag									*
 *																					*
 * Theory			=> This is an amazing effect, very funny, and very simple to	*
 *					understand. You just need understand how sin and cos works		*
 *																					*/
HRESULT __stdcall GPX_Waves (HDC	PicDestDC, 
							 HDC	PicSrcDC,
							 int	Amplitude, 
							 int	Frequency, 
							 char	FillSides, 
							 char	Direction,
							 int	*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (Amplitude < 0)
		Amplitude = 0;
	if (Frequency < 0)
		Frequency = 0;

	int h, w;
	if (Direction)		// Horizontal
	{
		int tx;
		for (h = 0; h < Height; h++)
		{
			tx = (int)(Amplitude * sin ((Frequency * 2) * h * (PI / 180)));
			::BitBlt (PicDestDC, tx, h, Width, 1, PicSrcDC, 0, h, SRCCOPY);
			if (FillSides)
			{
				::BitBlt (PicDestDC, 0, h, tx, 1, PicSrcDC, Width - tx, h, SRCCOPY);
				::BitBlt (PicDestDC, Width + tx, h, Width - (Width - 2 * Amplitude + tx), 
					1, PicSrcDC, 0, h, SRCCOPY);
			}
		}
	}
	else
	{
		int ty;
		for (w = 0; w < Width; w++)
		{
			ty = (int)(Amplitude * sin ((Frequency * 2) * w * (PI / 180)));
			::BitBlt (PicDestDC, w, ty, 1, Height, PicSrcDC, w, 0, SRCCOPY);
			if (FillSides)
			{
				::BitBlt (PicDestDC, w, 0, 1, ty, PicSrcDC, w, Height - ty, SRCCOPY);
				::BitBlt (PicDestDC, w, Height + ty, 1, Height - (Height - 2 * Amplitude + ty),
					PicSrcDC, w, 0, SRCCOPY);
			}
		}
	}
	::DeleteObject (PicSrcHwnd);

	*Response = 1;
	return (S_OK);
}

/* Function to apply the BlockWaves effect 											*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Amplitude		=> Sinoidal maximum height										*
 * Frequency		=> Frequency value												*
 * Mode				=> If Mode is an even number, Mode2 will be applied, else, Mode1*
 *					will be applied.												*
 *																					*
 * Theory			=> This is an amazing effect, very funny when amplitude and		*
 *					frequency are small values.										*
 *																					*/
HRESULT __stdcall GPX_BlockWaves (HDC		PicDestDC, 
								  HDC		PicSrcDC, 
								  short		Amplitude,
								  short		Frequency,
								  short		Mode,
								  int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (Amplitude < 0)
		Amplitude = 0;
	if (Frequency < 0)
		Frequency = 0;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];
	LPBYTE NewBits = new BYTE[BitCount];

	Result = ::GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0, nw, nh;
		double Radius;
		for (int w = 0; w < Width; w++)
			for (int h = 0; h < Height; h++)
			{
				i = h * LineWidth + 3 * w;
				nw = Width / 2 - w;
				nh = Height / 2 - h;
				Radius = sqrt (nw * nw + nh * nh);
				
				if (Mode % 2 == 0)
				{
					nw = (int)(w + Amplitude * sin (Frequency * nw * (PI / 180)));
					nh = (int)(h + Amplitude * cos (Frequency * nh * (PI / 180)));
				}
				else
				{
					nw = (int)(w + Amplitude * sin (Frequency * w * (PI / 180)));
					nh = (int)(h + Amplitude * cos (Frequency * h * (PI / 180)));
				}

				nw = (nw < 0) ? 0 : ((nw >= Width) ? Width - 1 : nw);
				nh = (nh < 0) ? 0 : ((nh >= Height) ? Height - 1 : nh);
				j = nh * LineWidth + 3 * nw;
				NewBits[i+2] = Bits[j+2];
				NewBits[i+1] = Bits[j+1];
				NewBits[ i ] = Bits[ j ];
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, NewBits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;
		delete [] NewBits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the DetectBorders effect 										*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Border			=> Border sensibility											*
 * ColorBorder		=> Border color													*
 * BGColor			=> BackGround color												*
 *																					*
 * Theory			=> This effect analize and plot borders. Use only 2 colors		*
 *																					*/
HRESULT __stdcall GPX_DetectBorders (HDC		PicDestDC, 
									 HDC		PicSrcDC,
									 int		Border, 
									 int		ColorBorder, 
									 int		BGColor,
									 int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	BYTE BackR = (BGColor & 0x000000FF),
				  BackG = (BGColor & 0x0000FF00) >> 8,
			      BackB = (BGColor & 0x00FF0000) >> 16;

	BYTE BorderR = (ColorBorder & 0x000000FF),
				  BorderG = (ColorBorder & 0x0000FF00) >> 8,
			      BorderB = (ColorBorder & 0x00FF0000) >> 16;

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	Border = LimitValues (Border);
	
	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0, color_1, color_2;
		BYTE GR, GG, GB, Gray;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				for (int k = 0; k <= 2; k++)
				{
					i = h * LineWidth + 3 * w;
					j = h * LineWidth + 3 * (w + Lim_Max (w, 1, Width));
					color_1 = (int)((Bits[i+k] - Bits[j+k]) * (Bits[i+k] - Bits[j+k]));
					j = (h + Lim_Max (h, 1, Height)) * LineWidth + 3 * w;
					color_2 = (int)((Bits[i+k] - Bits[j+k]) * (Bits[i+k] - Bits[j+k]));

					switch (k)
					{
						case 0:
							GB = LimitValues ((int)(sqrt ((color_1 + color_2) << 1)));
							break;
						case 1:
							GG = LimitValues ((int)(sqrt ((color_1 + color_2) << 1)));
							break;
						case 2:
							GR = LimitValues ((int)(sqrt ((color_1 + color_2) << 1)));
							break;
					}
				}

				Gray = (GR + GG + GB) / 3;
				if (Gray > Border)
				{
					Bits[i+2] = BorderR;
					Bits[i+1] = BorderG;
					Bits[ i ] = BorderB;
				}
				else
				{
					Bits[i+2] = BackR;
					Bits[i+1] = BackG;
					Bits[ i ] = BackB;
				}
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the Blur effect												*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 *																					*
 * Theory			=> This is a simple blur function, very easy to undestand.		*
 *					We get a pixel, this pixel is the center of an 3x3 matrix, all	*
 *					the neighboor pixels will be added and divided by 9 (total of	*
 *					pixels taken). This result will be applied over the center pixel*
 *																					*/
HRESULT __stdcall GPX_Blur (HDC		PicDestDC, 
							HDC		PicSrcDC,
							int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0, k = 0;
		for (int h = 1; h < Height - 1; h++)
			for (int w = 1; w < Width - 1; w++)
			{
				i = h * LineWidth + 3 * w;
				j = (h + 1) * LineWidth + 3 * w;
				k = (h - 1) * LineWidth + 3 * w;

				Bits[i+2] = (Bits[i-1] + Bits[j-1] + Bits[k-1] +
							 Bits[i+2] + Bits[j+2] + Bits[k+2] +
							 Bits[i+5] + Bits[j+5] + Bits[k+5]) / 9;
				Bits[i+1] = (Bits[i-2] + Bits[j-2] + Bits[k-2] +
							 Bits[i+1] + Bits[j+1] + Bits[k+1] +
							 Bits[i+4] + Bits[j+4] + Bits[k+4]) / 9;
				Bits[ i ] = (Bits[i-3] + Bits[j-3] + Bits[k-3] +
							 Bits[ i ] + Bits[ j ] + Bits[ k ] +
							 Bits[i+3] + Bits[j+3] + Bits[k+3]) / 9;
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the Relief effect												*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 *																					*
 * Theory			=> Mix emboss and neon effects, easy to understand, the			*
 *					difference between 2 pixels will be the key of this function	*
 *					If this difference is considerably, will create a color border.	*
 *																					*/
HRESULT __stdcall GPX_Relief (HDC		PicDestDC, 
							  HDC		PicSrcDC,
							  int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				j = (h + Lim_Max (h, 2, Height)) * LineWidth + 3 * (w + Lim_Max (w, 2, Width));
				
				Bits[i+2] = LimitValues ((int)((Bits[i+2] - Bits[j+2]) + 128));
				Bits[i+1] = LimitValues ((int)((Bits[i+1] - Bits[j+1]) + 128));
				Bits[ i ] = LimitValues ((int)((Bits[ i ] - Bits[ j ]) + 128));
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the Saturation effect											*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Saturation		=> Saturation value												*
 *																					*
 * Theory			=> This is a great color adjustment. If Saturation is less than	*
 *					255, all the colors will be approaching the gray tone. Else,	*
 *					all the colors will be being more intense.						*
 *																					*/
HRESULT __stdcall GPX_Saturation (HDC		PicDestDC, 
								  HDC		PicSrcDC,
								  int		Saturation,
								  int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (Saturation < -255)
		Saturation = -255;
	if (Saturation > 512)
		Saturation = 512;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, Gray;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				Gray = (Bits[i+2] + Bits[i+1] + Bits[i]) / 3;
				Bits[i+2] = LimitValues (ShadeColors (Gray, (int)Bits[i+2], 255 + Saturation));
				Bits[i+1] = LimitValues (ShadeColors (Gray, (int)Bits[i+1], 255 + Saturation));
				Bits[ i ] = LimitValues (ShadeColors (Gray, (int)Bits[ i ], 255 + Saturation));
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the FindEdges effect											*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Intensity		=> Intensity value												*
 * BW				=> Border width													*
 *																					*
 * Theory			=> Wow, another Photoshop filter (FindEdges). Do you understand	*
 *					Neon effect ? This is the same engine, but is inversed with		*
 *					255 - color. Easy, hun?											*
 *																					*/
HRESULT __stdcall GPX_FindEdges (HDC		PicDestDC, 
								 HDC		PicSrcDC,
								 short		Intensity, 
								 short		BW,
								 int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	Intensity = (Intensity < 0) ? 0 : (Intensity > 5) ? 5 : Intensity;
	BW = (BW < 1) ? 1 : (BW > 5) ? 5 : BW;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0, color_1, color_2;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
				for (int k = 0; k <= 2; k++)
				{
					i = h * LineWidth + 3 * w;
					j = h * LineWidth + 3 * (w + Lim_Max (w, BW, Width));
					color_1 = (int)((Bits[i+k] - Bits[j+k]) * (Bits[i+k] - Bits[j+k]));
					j = (h + Lim_Max (h, BW, Height)) * LineWidth + 3 * w;
					color_2 = (int)((Bits[i+k] - Bits[j+k]) * (Bits[i+k] - Bits[j+k]));
					Bits[i+k] = 255 - LimitValues ((int)(sqrt ((color_1 + color_2) << Intensity)));
				}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to calc the buffer size for AsciiMorph, just calc, ok?					*
 *																					*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 *																					*
 * Theory			=> A terminal character has aprox. 8 pixels (h) and 6 pixels (w)*
 *					With these values we can get the alloc size for AsciiBuffer		*
 *																					*/
HRESULT __stdcall GPX_AllocBufferSize (HDC		PicSrcDC,
									   int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	Result = ((Width / 3) + 2) * Height;
	::DeleteObject (PicSrcHwnd);

	*Response = (int)Result;
	return (S_OK);
}

/* Function to apply the AsciiMorph effect											*
 *																					*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * sBuffer			=> OutBuffer for the string										*
 *																					*
 * Theory			=> First, you need to create a char table sorted by darken		*
 *					After, you get the following formula: 255 / chars_number.		*
 *					i.e. if I create a table with 10 chars, the coeff is 25.5		*
 *					Now, you find the gray tone and see what char is darken as the	*
 *					gray tone, after this, you add this character to a buffer.		*
 *																					*/
HRESULT __stdcall GPX_AsciiMorph (HDC		PicSrcDC,
								  LPTSTR	sBuffer,
								  int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	int GrayColor;
	float color_table;

	// Static table for AsciiMorph
	//static string table[] = {"#", "W", "M", "B", "R", "X", "V", "Y", "I",  
	//				         "t", "i", "+", "=", ";", ":", ",", ".", " "};
	static char table[] = {'#', 'W', 'M', 'B', 'R', 'X', 'V', 'Y', 'I',  
					       't', 'i', '+', '=', ';', ':', ',', '.', ' '};

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);

	_tcscpy (sBuffer, "");							// Erase buffer
	if (Result)
	{
		int i = 0, pos = 0;
		for (int h = 0; h < Height; h += 4)
		{
			for (int w = 0; w < Width; w += 3)
			{
				i = (Height - h - 1) * LineWidth + 3 * w;
				GrayColor = (Bits[i] + Bits[i+1] + Bits[i+2]) / 3;
				color_table = (float)(GrayColor / 14.22);
				color_table = (color_table > 17) ? 17 : (color_table < 0) ? 0 : color_table;
				//_tcscat (sBuffer, table[(int)color_table]);
				sBuffer[pos++] = table[(int)color_table];
			}
			//_tcscat (sBuffer, "\x0d");
			//_tcscat (sBuffer, "\x0a");
			sBuffer[pos++] = '\x0d';
			sBuffer[pos++] = '\x0a';
		}

		sBuffer[pos] = '\0';

		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to Adjust the color hue													*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Hue				=> Hue value													*
 *																					*
 * Theory			=> Ok, you saw in PSC some Hue functions and is very different	*
 *					with this function, I believe in you. I was trying to understand*
 *					this adjustment and I saw that are many swap. I just apply this	*
 *					at my function.													*
 *																					*/
HRESULT __stdcall GPX_Hue (HDC		PicDestDC, 
						   HDC		PicSrcDC, 
						   int		Hue,
						   int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (Hue > 350)
		Hue = 350;
	if (Hue < 0)
		Hue = 0;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0, table;
		BYTE temp;
		table = Hue / 50;
		if (table > 6)
			table = 6;
		for (int w = 0; w < Width; w++)
			for (int h = 0; h < Height; h++)
			{
				i = h * LineWidth + 3 * w;
				switch (table)
				{
					case 0:
						Bits[i+1] = (BYTE)ShadeColors ((int)Bits[i+1], (int)Bits[i+2], (int)(Hue * 5.12));
						break;
					case 1:
						Bits[i+1] = Bits[i+2];
						Bits[i+2] = (BYTE)ShadeColors ((int)Bits[i+2], (int)Bits[ i ], (int)((Hue -  50) * 5.12));
						break;
					case 2:
						Bits[i+1] = Bits[i+2];
						Bits[i+2] = Bits[ i ];
						Bits[ i ] = (BYTE)ShadeColors ((int)Bits[ i ], (int)Bits[i+1], (int)((Hue - 100) * 5.12));
						break;
					case 3:
						temp = Bits[ i ];
						Bits[i+1] = Bits[ i ] = Bits[i+2];
						Bits[i+2] = temp;
						Bits[i+1] = (BYTE)ShadeColors ((int)Bits[i+1], (int)Bits[i+2], (int)((Hue - 150) * 5.12));
						break;
					case 4:
						temp = Bits[ i ];
						Bits[ i ] = Bits[i+2];
						Bits[i+1] = Bits[i+2] = temp;
						Bits[i+2] = (BYTE)ShadeColors ((int)Bits[i+2], (int)Bits[ i ], (int)((Hue - 200) * 5.12));
						break;
					case 5:
						Bits[i+1] = Bits[ i ];
						Bits[ i ] = Bits[i+2];
						Bits[ i ] = (BYTE)ShadeColors ((int)Bits[ i ], (int)Bits[i+1], (int)((Hue - 250) * 5.11));
						break;
					case 6:
						temp = Bits[i+1];
						Bits[i+1] = Bits[ i ];
						Bits[i+1] = (BYTE)ShadeColors ((int)Bits[i+1], (int)temp, (int)((Hue - 300) * 5.11));
						break;
				}
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the AlphaBlend effect											*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC_1		=> Source #1 PictureBox's Device Context						*
 * PicSrcDC_2		=> Source #2 PictureBox's Device Context						*
 * Alpha			=> Alpha value													*
 *																					*
 * Theory			=> This is a great effect and very easy to undestand, with		*
 *					Alpha value, we can get the proportional color between two		*
 *					pictures.														*
 *																					*/
HRESULT __stdcall GPX_AlphaBlend (HDC		PicDestDC,
								  HDC		PicSrcDC_1,
								  HDC		PicSrcDC_2, 
								  int		Alpha,
								  int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width = 0, Height = 0; 
	int SrcWidth = 0, SrcHeight = 0;
	int DestWidth = 0, DestHeight = 0;

	BITMAPINFO infod, infos;
	HBITMAP	PicSrcHwnd_1 = GetBitmapHandle (PicSrcDC_1, &infos, &SrcWidth, &SrcHeight);
	HBITMAP	PicSrcHwnd_2 = GetBitmapHandle (PicSrcDC_2, &infod, &DestWidth, &DestHeight);

	if (! (PicSrcHwnd_1 && PicSrcHwnd_2))
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	Alpha = (Alpha > 255) ? 255 : (Alpha < 0) ? 0 : Alpha;

	Width = (SrcWidth > DestWidth) ? DestWidth : SrcWidth;
	Height = (SrcHeight > DestHeight) ? DestHeight : SrcHeight;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE DestBits = new BYTE[BitCount];
	LPBYTE  SrcBits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC_2, PicSrcHwnd_2, 0, Height, SrcBits, &infos, DIB_RGB_COLORS);
	Result = GetDIBits (PicSrcDC_1, PicSrcHwnd_1, 0, Height, DestBits, &infod, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				DestBits[i+2] = (BYTE)ShadeColors ((int)DestBits[i+2], (int)SrcBits[i+2], Alpha);
				DestBits[i+1] = (BYTE)ShadeColors ((int)DestBits[i+1], (int)SrcBits[i+1], Alpha);
				DestBits[ i ] = (BYTE)ShadeColors ((int)DestBits[ i ], (int)SrcBits[ i ], Alpha);
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, DestBits, &infod, 0);
		
		::DeleteObject (PicSrcHwnd_1);
		::DeleteObject (PicSrcHwnd_2);
		delete [] DestBits;
		delete [] SrcBits;
		
		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the Make3DEffect effect										*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Normal			=> Normal value													*
 *																					*
 * Theory			=> A weird effect, its very cool with some pictures. The engine	*
 *					isn't hard to understand. We create lines depending the color	*
 *																					*/
HRESULT __stdcall GPX_Make3DEffect (HDC		PicDestDC, 
									HDC		PicSrcDC,
									int		Normal,
									int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	Normal = (Normal > 50) ? 50 : (Normal < 1) ? 1 : Normal;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];
	LPBYTE NewBits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j, Step;
		for (int w = 0; w < Width; w++)
			for (int h = 0; h < Height; h++)
			{
				i = (Height - h - 1) * LineWidth + 3 * w;
				Step = (int)(((Bits[i+2] + Bits[i+1] + Bits[i]) / 3) / Normal);
				if (h - Step > 0)
				{
					for (int y = 0; y < Step; y++)
					{
						j = (Height - (h - Lim_Max (h, y, Height)) - 1) * LineWidth + 3 * w;
						NewBits[j+2] = Bits[i+2];
						NewBits[j+1] = Bits[i+1];
						NewBits[ j ] = Bits[ i ];
					}
				}
				else
				{
					NewBits[i+2] = Bits[i+2];
					NewBits[i+1] = Bits[i+1];
					NewBits[ i ] = Bits[ i ];
				}

			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, NewBits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;
		delete [] NewBits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the FourCorners effect											*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 *																					*
 * Theory			=> This is an amazing function, you've never seen this before.	*
 *					I was testing some trigonometric functions, and I saw that if	*
 *					I multiply the angle by 2, the result is an image like this		*
 *					If we multiply by 3, we can create the SixCorners effect.		*
 *																					*/
HRESULT __stdcall GPX_FourCorners (HDC		PicDestDC, 
								   HDC		PicSrcDC,
								   int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];
	LPBYTE NewBits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0, nw, nh;
		double Angle, Radius, rNew;
		double Radmax = sqrt (Width * Width + Height * Height) / 2;
		for (int w = 0; w < Width; w++)
			for (int h = 0; h < Height; h++)
			{
				nw = Width / 2 - w;
				nh = Height / 2 - h;
				Radius = sqrt (nw * nw + nh * nh);
				if (Radius < Radmax)
				{
					Angle = atan2 (nh, nw) * 2;
					rNew = Radius * Radius / Radmax;
					nw = (int)(Width / 2 + rNew * cos (Angle));
					nh = (int)(Height / 2 - rNew * sin (Angle));
					nw = (nw < 0) ? 0 : ((nw >= Width) ? Width - 1 : nw);
					nh = (nh < 0) ? 0 : ((nh >= Height) ? Height - 1 : nh);
					i = h * LineWidth + 3 * w;
					j = nh * LineWidth + 3 * nw;
					NewBits[i+2] = Bits[j+2];
					NewBits[i+1] = Bits[j+1];
					NewBits[ i ] = Bits[ j ];
				}
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, NewBits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;
		delete [] NewBits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the Caricature effect											*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 *																					*
 * Theory			=> Caricature is a nice effect, the theory is similar to FishEye*
 *					but inversed.													*
 *																					*/
HRESULT __stdcall GPX_Caricature (HDC		PicDestDC, 
								  HDC		PicSrcDC,
								  int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];
	LPBYTE NewBits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0, nw, nh;
		double Angle, Radius, rNew;
		double Radmax = sqrt (Width * Width + Height * Height) / 2;
		for (int w = 0; w < Width; w++)
			for (int h = 0; h < Height; h++)
			{
				nw = Width / 2 - w;
				nh = Height / 2 - h;
				Radius = sqrt (nw * nw + nh * nh);
				if (Radius < Radmax)
				{
					Angle = atan2 (nh, nw);
					rNew = sqrt (Radius * Radmax);
					nw = (int)(Width / 2 + rNew * cos (Angle));
					nh = (int)(Height / 2 - rNew * sin (Angle));
					nw = (nw < 0) ? 0 : ((nw >= Width) ? Width - 1 : nw);
					nh = (nh < 0) ? 0 : ((nh >= Height) ? Height - 1 : nh);
					i = h * LineWidth + 3 * (Width - w - 1);
					j = nh * LineWidth + 3 * nw;
					NewBits[i+2] = Bits[j+2];
					NewBits[i+1] = Bits[j+1];
					NewBits[ i ] = Bits[ j ];
				}
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, NewBits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;
		delete [] NewBits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the Tile effect												*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * WSize			=> Width														*
 * HSize			=> Height														*
 * Random			=> Maximum random value											*
 *																					*
 * Theory			=> Similar to Tile effect from Photoshop and very easy to		*
 *					understand. We get a rectangular area using WSize and HSize and	*
 *					replace in a position with a random distance from the original	*
 *					position.														*
 *																					*/
HRESULT __stdcall GPX_Tile (HDC		PicDestDC, 
							HDC		PicSrcDC,
							int		WSize, 
							int		HSize, 
							int		Random,
							int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (WSize < 1)
		WSize = 1;
	if (HSize < 1)
		HSize = 1;
	if (Random < 1)
		Random = 1;
	
	srand ((UINT) GetTickCount());
	int tx, ty, h, w;
	for (h = 0; h < Height; h += HSize)
		for (w = 0; w < Width; w += WSize)
		{
			tx = (int)(rand() % Random) - (Random / 2);
			ty = (int)(rand() % Random) - (Random / 2);
			::BitBlt (PicDestDC, w + tx, h + ty, WSize, HSize, PicSrcDC, w, h, SRCCOPY);
		}
	::DeleteObject (PicSrcHwnd);

	*Response = 1;
	return (S_OK);
}

/* Function to apply the Roll effect												*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 *																					*
 * Theory			=> Roll function isn't very interesting, but has some essential	*
 *					trigonometric functions.										*
 *																					*/
HRESULT __stdcall GPX_Roll (HDC		PicDestDC, 
							HDC		PicSrcDC,
							int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];
	LPBYTE NewBits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0, nw, nh;
		double Angle, Radius;
		double Radmax = sqrt (Width * Width + Height * Height) / 2;
		for (int w = 0; w < Width; w++)
			for (int h = 0; h < Height; h++)
			{
				nw = Width / 2 - w;
				nh = Height / 2 - h;
				Radius = sqrt (nw * nw + nh * nh);
				if (Radius < Radmax)
				{
					Angle = atan2 (nh, nw) / 2;
					nw = (int)(Width / 2 + Radius * cos (Angle));
					nh = (int)(Height / 2 - Radius * sin (Angle));
					nw = (nw < 0) ? 0 : ((nw >= Width) ? Width - 1 : nw);
					nh = (nh < 0) ? 0 : ((nh >= Height) ? Height - 1 : nh);
					i = h * LineWidth + 3 * (Width - w - 1);
					j = nh * LineWidth + 3 * nw;
					NewBits[i+2] = Bits[j+2];
					NewBits[i+1] = Bits[j+1];
					NewBits[ i ] = Bits[ j ];
				}
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, NewBits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;
		delete [] NewBits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the SmartBlur effect											*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Sensibility		=> SmartBlur sensibility										*
 *																					*
 * Theory			=> Similar to SmartBlur from Photoshop, this function has the	*
 *					same engine as Blur function, but, in a matrix with 3x3			*
 *					dimentions, we take only colors that pass by sensibility filter	*
 *					The result is a clean image, not totally blurred, but a image	*
 *					with correction between pixels.	A great effect.					*
 *																					*/
HRESULT __stdcall GPX_SmartBlur (HDC		PicDestDC, 
								 HDC		PicSrcDC, 
								 int		Sensibility,
								 int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	Sensibility = LimitValues (Sensibility);

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];
	BYTE Temp[3][9];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0, k = 0, GrayCmp, Gray, add;
		for (int h = 1; h < Height - 1; h++)
			for (int w = 1; w < Width - 1; w++)
			{
				i = h * LineWidth + 3 * w;
				j = (h + 1) * LineWidth + 3 * w;
				k = (h - 1) * LineWidth + 3 * w;
				Gray = (Bits[i+2] + Bits[i+1] + Bits[i]) / 3;

				for (int y = 0; y < 3; y++)
				{
					add = y * 3;
					GrayCmp = (Bits[j+add-1] + Bits[j+add-2] + Bits[j+add-3]) / 3;
					if ((GrayCmp > Gray + Sensibility) || (GrayCmp < Gray - Sensibility))
					{
						Temp[0][y] = Bits[ i ];
						Temp[1][y] = Bits[i+1];
						Temp[2][y] = Bits[i+2];
					}
					else
					{
						Temp[0][y] = Bits[j+add-3];
						Temp[1][y] = Bits[j+add-2];
						Temp[2][y] = Bits[j+add-1];
					}
				}

				for (y = 0; y < 3; y++)
				{
					add = y * 3;
					GrayCmp = (Bits[i+add-1] + Bits[i+add-2] + Bits[i+add-3]) / 3;
					if ((GrayCmp > Gray + Sensibility) || (GrayCmp < Gray - Sensibility))
					{
						Temp[0][y+3] = Bits[ i ];
						Temp[1][y+3] = Bits[i+1];
						Temp[2][y+3] = Bits[i+2];
					}
					else
					{
						Temp[0][y+3] = Bits[i+add-3];
						Temp[1][y+3] = Bits[i+add-2];
						Temp[2][y+3] = Bits[i+add-1];
					}
				}

				for (y = 0; y < 3; y++)
				{
					add = y * 3;
					GrayCmp = (Bits[k+add-1] + Bits[k+add-2] + Bits[k+add-3]) / 3;
					if ((GrayCmp > Gray + Sensibility) || (GrayCmp < Gray - Sensibility))
					{
						Temp[0][y+6] = Bits[ i ];
						Temp[1][y+6] = Bits[i+1];
						Temp[2][y+6] = Bits[i+2];
					}
					else
					{
						Temp[0][y+6] = Bits[k+add-3];
						Temp[1][y+6] = Bits[k+add-2];
						Temp[2][y+6] = Bits[k+add-1];
					}
				}

				Bits[i+2] = (Temp[2][0] + Temp[2][1] + Temp[2][2] +
							 Temp[2][3] + Temp[2][4] + Temp[2][5] +
							 Temp[2][6] + Temp[2][7] + Temp[2][8]) / 9;
				Bits[i+1] = (Temp[1][0] + Temp[1][1] + Temp[1][2] +
							 Temp[1][3] + Temp[1][4] + Temp[1][5] +
							 Temp[1][6] + Temp[1][7] + Temp[1][8]) / 9;
				Bits[ i ] = (Temp[0][0] + Temp[0][1] + Temp[0][2] +
							 Temp[0][3] + Temp[0][4] + Temp[0][5] +
							 Temp[0][6] + Temp[0][7] + Temp[0][8]) / 9;
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the AdvancedBlur effect										*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Blur				=> Blur radius (up to 10)										*
 * Sense			=> SmartBlur sensibility										*
 * Smart			=> To apply or not smartblur									*
 *																					*
 * Theory			=> This is a great effect, and mix blur and smartblur in one	*
 *					function. The maximum blur matrix is a matrix with 21x21		*
 *					dimentions. You can use the same matrix to do a better smartblur*
 *																					*/
HRESULT __stdcall GPX_AdvancedBlur (HDC		PicDestDC, 
									HDC		PicSrcDC, 
									short	Blur, 
									short	Sense, 
									BOOL	Smart,
									int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (Blur < 1)
		return (S_OK);
	if (Blur > 10)
		Blur = 10;

	int Orig = Blur; 
	Blur = Blur * 2 + 1;
	int Size = Blur * Blur;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];
	int SomaR = 0, SomaG = 0, SomaB = 0;

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i, j, GrayCmp, Gray;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				Gray = (Bits[i+2] + Bits[i+1] + Bits[i]) / 3;
				for (int a = -Orig; a <= Orig; a++)
					for (int b = -Orig; b <= Orig; b++)
					{
						j = (h + Lim_Max (h, a, Height)) * LineWidth + 3 * (w + Lim_Max (w, b, Width));
						if ((h + a < 0) || (w + b < 0))
							j = i;
						if (Smart)
						{
							GrayCmp = (Bits[j+2] + Bits[j+1] + Bits[j]) / 3;
							if ((GrayCmp > Gray + Sense) || (GrayCmp < Gray - Sense))
								j = i;
						}
						SomaR += Bits[j+2];
						SomaG += Bits[j+1];
						SomaB += Bits[ j ];
					}
				Bits[i+2] = SomaR / Size;
				Bits[i+1] = SomaG / Size;
				Bits[ i ] = SomaB / Size;
				SomaR = SomaG = SomaB = 0;
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the SoftnerBlur effect											*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 *																					*
 * Theory			=> An interesting blur-like function. In dark tones we apply a	*
 *					blur with 3x3 dimentions, in light tones, we apply a blur with	*
 *					5x5 dimentions. Easy, hun?										*
 *																					*/
HRESULT __stdcall GPX_SoftnerBlur (HDC		PicDestDC, 
								   HDC		PicSrcDC,
								   int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];
	int SomaR = 0, SomaG = 0, SomaB = 0;

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i, j, Gray;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				Gray = (Bits[i+2] + Bits[i+1] + Bits[i]) / 3;
				if (Gray > 127)
				{
					for (int a = -3; a <= 3; a++)
						for (int b = -3; b <= 3; b++)
						{
							j = (h + Lim_Max (h, a, Height)) * LineWidth + 3 * (w + Lim_Max (w, b, Width));
							if ((h + a < 0) || (w + b < 0))
								j = i;
							SomaR += Bits[j+2];
							SomaG += Bits[j+1];
							SomaB += Bits[ j ];
						}
					Bits[i+2] = SomaR / 49;
					Bits[i+1] = SomaG / 49;
					Bits[ i ] = SomaB / 49;
					SomaR = SomaG = SomaB = 0;
				}
				else
				{
					for (int a = -1; a <= 1; a++)
						for (int b = -1; b <= 1; b++)
						{
							j = (h + Lim_Max (h, a, Height)) * LineWidth + 3 * (w + Lim_Max (w, b, Width));
							if ((h + a < 0) || (w + b < 0))
								j = i;
							SomaR += Bits[j+2];
							SomaG += Bits[j+1];
							SomaB += Bits[ j ];
						}
					Bits[i+2] = SomaR / 9;
					Bits[i+1] = SomaG / 9;
					Bits[ i ] = SomaB / 9;
					SomaR = SomaG = SomaB = 0;
				}
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the MotionBlur effect											*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Angle			=> Angle direction (degrees)									*
 * Distance			=> The one-dimention blur value									*
 *																					*
 * Theory			=> Similar to MotionBlur from Photoshop, the engine is very		*
 *					simple to undertand, we take a pixel (duh!), with the angle we	*
 *					will taking near pixels. After this we blur (add and do a		*
 *					division).														*
 *																					*/
HRESULT __stdcall GPX_MotionBlur (HDC		PicDestDC, 
								  HDC		PicSrcDC,
								  double	Angle, 
								  int		Distance,
								  int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (Angle == 0)
		Angle = 0.0001;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];
	LPBYTE NewBits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j, R = 0, G = 0, B = 0, Size = Distance * 2 + 1;
		double nw, nh;
		for (double h = 0; h < Height; h++)
			for (double w = 0; w < Width; w++)
			{
				i = (int)h * LineWidth + 3 * (int)w;
				for (double a = -Distance; a <= Distance; a++)
				{
					nw = (double)(w + a * cos ((2 * PI) / (360 / Angle)));
					nh = (double)(h + a * sin ((2 * PI) / (360 / Angle)));
					nw = (nw >= Width) ? Width - 1 : (nw < 0) ? 0 : nw;
					nh = (nh >= Height) ? Height - 1 : (nh < 0) ? 0 : nh;
					j = (int)nh * LineWidth + 3 * (int)nw;
					R += Bits[j+2];
					G += Bits[j+1];
					B += Bits[ j ];
				}

				NewBits[i+2] = R / Size;
				NewBits[i+1] = G / Size;
				NewBits[ i ] = B / Size;
				R = G = B = 0;
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, NewBits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;
		delete [] NewBits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the ColorBalance effect										*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * R				=> Red value													*
 * G				=> Green value													*
 * B				=> Blue Value													*
 *																					*
 * Theory			=> Similar to brightness function, but here we can choice what	*
 *					tone we will apply.												*
 *																					*/
HRESULT __stdcall GPX_ColorBalance (HDC		PicDestDC, 
									HDC		PicSrcDC,
									short	R, 
									short	G, 
									short	B,
									int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				Bits[i+2] = LimitValues ((int)(Bits[i+2] + R));
				Bits[i+1] = LimitValues ((int)(Bits[i+1] + G));
				Bits[ i ] = LimitValues ((int)(Bits[ i ] + B));
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the Fragment effect											*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Distance			=> Distance between layers (from origin)						*
 *																					*
 * Theory			=> Similar to Fragment effect from Photoshop. We create 4 layers*
 *					each one has the same distance from the origin, but have		*
 *					different positions (top, botton, left and right), with these 4	*
 *					layers, we join all the pixels.									*
 *																					*/
HRESULT __stdcall GPX_Fragment (HDC		PicDestDC, 
								HDC		PicSrcDC, 
								int		Distance,
								int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];
	LPBYTE Layer1 = new BYTE[BitCount];
	LPBYTE Layer2 = new BYTE[BitCount];
	LPBYTE Layer3 = new BYTE[BitCount];
	LPBYTE Layer4 = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0, nw, nh;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;

				nh = (h + Distance >= Height) ? Height - 1 : h + Distance;
				j = nh * LineWidth + 3 * w;
				Layer1[i+2] = Bits[j+2];
				Layer1[i+1] = Bits[j+1];
				Layer1[ i ] = Bits[ j ];

				nh = (h - Distance < 0) ? 0 : h - Distance;
				j = nh * LineWidth + 3 * w;
				Layer2[i+2] = Bits[j+2];
				Layer2[i+1] = Bits[j+1];
				Layer2[ i ] = Bits[ j ];

				nw = (w + Distance >= Width) ? Width - 1 : w + Distance;
				j = h * LineWidth + 3 * nw;
				Layer3[i+2] = Bits[j+2];
				Layer3[i+1] = Bits[j+1];
				Layer3[ i ] = Bits[ j ];

				nw = (w - Distance < 0) ? 0 : w - Distance;
				j = h * LineWidth + 3 * nw;
				Layer4[i+2] = Bits[j+2];
				Layer4[i+1] = Bits[j+1];
				Layer4[ i ] = Bits[ j ];
			}

		for (h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				Bits[i+2] = (Layer1[i+2] + Layer2[i+2] + Layer3[i+2] + Layer4[i+2]) / 4;
				Bits[i+1] = (Layer1[i+1] + Layer2[i+1] + Layer3[i+1] + Layer4[i+1]) / 4;
				Bits[ i ] = (Layer1[ i ] + Layer2[ i ] + Layer3[ i ] + Layer4[ i ]) / 4;
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;
		delete [] Layer1;
		delete [] Layer2;
		delete [] Layer3;
		delete [] Layer4;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the FarBlur effect												*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Distance			=> Distance value												*
 *																					*
 * Theory			=> This is an interesting effect, the blur is applied in that	*
 *					way: (the value "1" means pixel to be used in a blur calc, ok?) *
 *					e.g. With distance = 2	 _ _ _ _ _								*
 *											|1|1|1|1|1|								*
 *											|1|0|0|0|1|								*
 *											|1|0|C|0|1|								*
 *											|1|0|0|0|1|								*
 *											|1|1|1|1|1|								*
 *					We sum all the pixels with value = 1 and apply at the pixel with*
 *					the position "C".												*
 *																					*/
HRESULT __stdcall GPX_FarBlur (HDC		PicDestDC, 
							   HDC		PicSrcDC, 
							   int		Distance,
							   int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (Distance < 1)
		Distance = 1;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];
	LPBYTE NewBits = new BYTE[BitCount];
	int SomaR = 0, SomaG = 0, SomaB = 0;

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i, j, nw, nh, counter = 0;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				for (int a = -Distance; a <= Distance; a++)
					for (int b = -Distance; b <= Distance; b++)
					{
						if ((a == -Distance) || (a == Distance))
						{
							nw = (w + b < 0) ? 0 : (w + b >= Width) ? Width - 1 : w + b;
							nh = (h + a < 0) ? 0 : (h + a >= Height) ? Height - 1 : h + a;
							j = nh * LineWidth + 3 * nw;
							SomaR += Bits[j+2];
							SomaG += Bits[j+1];
							SomaB += Bits[ j ];
							counter++;
						}
						else
						{
							nw = (w - Distance < 0) ? 0 : (w - Distance >= Width) ? Width - 1 : w - Distance;
							nh = (h + a < 0) ? 0 : (h + a >= Height) ? Height - 1 : h + a;
							j = nh * LineWidth + 3 * nw;
							SomaR += Bits[j+2];
							SomaG += Bits[j+1];
							SomaB += Bits[ j ];

							nw = (w + Distance < 0) ? 0 : (w + Distance >= Width) ? Width - 1 : w + Distance;
							nh = (h + a < 0) ? 0 : (h + a >= Height) ? Height - 1 : h + a;
							j = nh * LineWidth + 3 * nw;
							SomaR += Bits[j+2];
							SomaG += Bits[j+1];
							SomaB += Bits[ j ];
							
							counter += 2;
							b = Distance;
						}
					}
				
				i = h * LineWidth + 3 * w;
				NewBits[i+2] = SomaR / counter;
				NewBits[i+1] = SomaG / counter;
				NewBits[ i ] = SomaB / counter;
				SomaR = SomaG = SomaB = counter = 0;							
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, NewBits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;
		delete [] NewBits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the RadialBlur effect											*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Distance			=> Distance value												*
 *																					*
 * Theory			=> Similar to RadialBlur from Photoshop, its an amazing effect	*
 *					Very easy to understand but a little hard to implement.			*
 *					We have all the image and find the center pixel. Now, we analize*
 *					all the pixels and calc the radius from the center and find the	*
 *					angle. After this, we sum this pixel with others with the same	*
 *					radius, but different angles. Here I'm using degrees angles.	*
 *																					*/
HRESULT __stdcall GPX_RadialBlur (HDC		PicDestDC, 
								  HDC		PicSrcDC,
								  short		Distance,
								  int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (Distance == 0)
		return (S_OK);
	if (Distance < 1)
		Distance = 1;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];
	LPBYTE NewBits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0, nw, nh, SomaR = 0, SomaG = 0, SomaB = 0, counter = 0;
		double AngleRad, Radius, Angle;
		for (int w = 0; w < Width; w++)
			for (int h = 0; h < Height; h++)
			{
				nw = Width / 2 - w;
				nh = Height / 2 - h;
				Radius = sqrt (nw * nw + nh * nh);
				AngleRad = atan2 (nh, nw);
				for (int a = -Distance; a <= Distance; a++)
				{
					Angle = AngleRad + (a * ANGLE_PERCENTAGE);
					nw = (int)(Width / 2 - Radius * cos (Angle));
					nh = (int)(Height / 2 - Radius * sin (Angle));
					nw = (nw < 0) ? 0 : ((nw >= Width) ? Width - 1 : nw);
					nh = (nh < 0) ? 0 : ((nh >= Height) ? Height - 1 : nh);
					j = nh * LineWidth + 3 * nw;
					SomaR += Bits[j+2];
					SomaG += Bits[j+1];
					SomaB += Bits[ j ];
					counter++;
				}

				i = h * LineWidth + 3 * w;
				NewBits[i+2] = SomaR / counter;
				NewBits[i+1] = SomaG / counter;
				NewBits[ i ] = SomaB / counter;
				SomaR = SomaG = SomaB = counter = 0;
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, NewBits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;
		delete [] NewBits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the ZoomBlur effect											*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Distance			=> Distance value												*
 *																					*
 * Theory			=> Here we have a effect similar to RadialBlur mode Zoom from	*
 *					Photoshop. The theory is very similar to RadialBlur, but has one*
 *					difference. Instead we use pixels with the same radius and		*
 *					near angles, we take pixels with the same angle but near radius	*
 *					This radius is always from the center to out of the image, we	*
 *					calc a proportional radius from the center.						*
 *																					*/
HRESULT __stdcall GPX_ZoomBlur (HDC		PicDestDC, 
								HDC		PicSrcDC, 
								short	Distance,
								int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (Distance == 0)
		return (S_OK);
	if (Distance < 1)
		Distance = 1;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];
	LPBYTE NewBits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0, nw, nh, SomaR = 0, SomaG = 0, SomaB = 0;
		double AngleRad, Radius, dNew, Radmax = sqrt (Height * Height + Width * Width);
		for (int w = 0; w < Width; w++)
			for (int h = 0; h < Height; h++)
			{
				nw = Width / 2 - w;
				nh = Height / 2 - h;
				Radius = sqrt (nw * nw + nh * nh);
				AngleRad = atan2 (nh, nw);
				dNew = (Radius * Distance) / Radmax;
				for (int a = 0; a < dNew; a++)
				{
					nw = (int)(Width / 2 - (Radius - a) * cos (AngleRad));
					nh = (int)(Height / 2 - (Radius - a) * sin (AngleRad));
					nw = (nw < 0) ? 0 : ((nw >= Width) ? Width - 1 : nw);
					nh = (nh < 0) ? 0 : ((nh >= Height) ? Height - 1 : nh);
					j = nh * LineWidth + 3 * nw;
					SomaR += Bits[j+2];
					SomaG += Bits[j+1];
					SomaB += Bits[ j ];
				}

				i = h * LineWidth + 3 * w;
				if (dNew)
				{
					NewBits[i+2] = SomaR / (int)(dNew + 1);
					NewBits[i+1] = SomaG / (int)(dNew + 1);
					NewBits[ i ] = SomaB / (int)(dNew + 1);
				}
				else
				{
					NewBits[i+2] = Bits[i+2];
					NewBits[i+1] = Bits[i+1];
					NewBits[ i ] = Bits[ i ];
				}
				SomaR = SomaG = SomaB = 0;
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, NewBits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;
		delete [] NewBits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the WebColors effect											*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 *																					*
 * Theory			=> This is an easy function to understand and very easy to		*
 *					implement. We get the image and the result is the image with	*
 *					only 216 colors.												*
 *																					*/
HRESULT __stdcall GPX_WebColors (HDC		PicDestDC, 
								 HDC		PicSrcDC,
								 int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0;
		BYTE Color;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				for (int k = 0; k < 3; k++)
				{
					Color = Bits[i+k];
					if (Color < 26) Bits[i+k] = 0;
					else if (Color <  77) Bits[i+k] = 51;
					else if (Color < 128) Bits[i+k] = 102;
					else if (Color < 179) Bits[i+k] = 153;
					else if (Color < 230) Bits[i+k] = 204;
					else Bits[i+k] = 255;
				}
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the Fog effect													*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Fog				=> Fog value													*
 *																					*
 * Theory			=> Think that a fog has the more grayest color (RGB = 127, 127,	*
 *					127), now, we adjust the pixel to approach this value.			*
 *																					*/
HRESULT __stdcall GPX_Fog (HDC		PicDestDC, 
						   HDC		PicSrcDC,
						   int		Fog,
						   int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				if (Bits[i+2] > 127)
				{
					Bits[i+2] = LimitValues (Bits[i+2] - Fog);
					Bits[i+2] = (Bits[i+2] <= 127) ? 127 : Bits[i+2];
				}
				else
				{
					Bits[i+2] = LimitValues (Bits[i+2] + Fog);
					Bits[i+2] = (Bits[i+2] > 127) ? 127 : Bits[i+2];
				}

				if (Bits[i+1] > 127)
				{
					Bits[i+1] = LimitValues (Bits[i+1] - Fog);
					Bits[i+1] = (Bits[i+1] <= 127) ? 127 : Bits[i+1];
				}
				else
				{
					Bits[i+1] = LimitValues (Bits[i+1] + Fog);
					Bits[i+1] = (Bits[i+1] > 127) ? 127 : Bits[i+1];
				}

				if (Bits[ i ] > 127)
				{
					Bits[ i ] = LimitValues (Bits[ i ] - Fog);
					Bits[ i ] = (Bits[ i ] <= 127) ? 127 : Bits[ i ];
				}
				else
				{
					Bits[ i ] = LimitValues (Bits[ i ] + Fog);
					Bits[ i ] = (Bits[ i ] > 127) ? 127 : Bits[ i ];
				}
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the MediumTones effect											*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Level			=> Histogram level adjust										*
 *																					*
 * Theory			=> This is an impressive effect and very easy to understand.	*
 *					Firstly, we sum ALL the pixels and divide by the total pixels.	*
 *					Now, we have the medium tone. Now, with level value, we calc	*
 *					the intensity.													*
 *																					*/
HRESULT __stdcall GPX_MediumTones (HDC		PicDestDC, 
								   HDC		PicSrcDC,
								   int		Level,
								   int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, HalfR, HalfG, HalfB;
		long RAdd = 0, GAdd = 0, BAdd = 0;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				RAdd += Bits[i+2];
				GAdd += Bits[i+1];
				BAdd += Bits[ i ];
			}

		HalfR = RAdd / (BitCount / 3) + 1;
		HalfG = GAdd / (BitCount / 3) + 1;
		HalfB = BAdd / (BitCount / 3) + 1;

		for (h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				Bits[i+2] = LimitValues ((Bits[i+2] * Level) / HalfR);
				Bits[i+1] = LimitValues ((Bits[i+1] * Level) / HalfG);
				Bits[ i ] = LimitValues ((Bits[ i ] * Level) / HalfB);
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the CircularWaves effect										*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Amplitude		=> Amplitude value												*
 * Frequency		=> Frequency value												*
 *																					*
 * Theory			=> Similar to Waves effect, but here I apply a senoidal function*
 *					with the angle point.											*
 *																					*/
HRESULT __stdcall GPX_CircularWaves (HDC		PicDestDC, 
									 HDC		PicSrcDC,
									 short		Amplitude,
									 short		Frequency,
									 int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (Amplitude < 0)
		Amplitude = 0;
	if (Frequency < 0)
		Frequency = 0;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];
	LPBYTE NewBits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0, nw, nh;
		double Radius;
		for (int w = 0; w < Width; w++)
			for (int h = 0; h < Height; h++)
			{
				i = h * LineWidth + 3 * w;
				nw = Width / 2 - w;
				nh = Height / 2 - h;
				Radius = sqrt (nw * nw + nh * nh);
				nw = (int)(w + Amplitude * sin (Frequency * Radius * (PI / 180)));
				nh = (int)(h + Amplitude * cos (Frequency * Radius * (PI / 180)));
				nw = (nw < 0) ? 0 : ((nw >= Width) ? Width - 1 : nw);
				nh = (nh < 0) ? 0 : ((nh >= Height) ? Height - 1 : nh);
				j = nh * LineWidth + 3 * nw;
				NewBits[i+2] = Bits[j+2];
				NewBits[i+1] = Bits[j+1];
				NewBits[ i ] = Bits[ j ];
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, NewBits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;
		delete [] NewBits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the CircularWavesEx effect 									*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Amplitude		=> Amplitude value												*
 * Frequency		=> Frequency value												*
 *																					*
 * Theory			=> Similar to CircularWaves effect, but the amplitude is		*
 *					proportional to radius.											*
 *																					*/
HRESULT __stdcall GPX_CircularWavesEx (HDC		PicDestDC, 
									   HDC		PicSrcDC, 
									   short	Amplitude,
									   short	Frequency,
									   int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (Amplitude < 0)
		Amplitude = 0;
	if (Frequency < 0)
		Frequency = 0;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];
	LPBYTE NewBits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0, nw, nh, NewAmp;
		double Radius, RadMax;
		for (int w = 0; w < Width; w++)
			for (int h = 0; h < Height; h++)
			{
				i = h * LineWidth + 3 * w;
				nw = Width / 2 - w;
				nh = Height / 2 - h;
				Radius = sqrt (nw * nw + nh * nh);
				RadMax = sqrt (Width * Width + Height * Height);
				NewAmp = (int)((Amplitude * Radius) / RadMax);
				nw = (int)(w + NewAmp * sin (Frequency * Radius * (PI / 180)));
				nh = (int)(h + NewAmp * cos (Frequency * Radius * (PI / 180)));
				nw = (nw < 0) ? 0 : ((nw >= Width) ? Width - 1 : nw);
				nh = (nh < 0) ? 0 : ((nh >= Height) ? Height - 1 : nh);
				j = nh * LineWidth + 3 * nw;
				NewBits[i+2] = Bits[j+2];
				NewBits[i+1] = Bits[j+1];
				NewBits[ i ] = Bits[ j ];
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, NewBits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;
		delete [] NewBits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the PolarCoordinates effect 									*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Flag				=> Not used yet													*
 *																					*
 * Theory			=> Similar to PolarCoordinates (Photoshop). We apply the polar	*
 *					transformation in a proportional (Height and Width) radius.		*
 *																					*/
HRESULT __stdcall GPX_PolarCoordinates (HDC		PicDestDC, 
										HDC		PicSrcDC,
										int		Flag,
										int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE    Bits = new BYTE[BitCount];
	LPBYTE NewBits = new BYTE[BitCount];
	LPBYTE   Flags = new BYTE[BitCount];
	ZeroMemory (Flags, BitCount);

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0;
		double k, l, nw, nh;
		double Angle, Radius, R;
		double m_w = Width / 2, m_h = Height / 2;

		for (int w = 0; w < Width; w++)
			for (int h = 0; h < Height; h++)
			{
				nw = m_w - w;
				nh = h - m_h;

				Radius = sqrt (nw * nw + nh * nh);
				Angle = atan2 (nh, nw);
				R = MaximumRadius (Height, Width, Angle);
				nh = (Radius * Height / R);
				nw = (Angle * Width / 6.2832);

				nh = Height - nh - 1;
				nw += m_w;

				nw = (nw < 0) ? 0 : ((nw >= Width) ? Width - 1 : nw);
				nh = (nh < 0) ? 0 : ((nh >= Height) ? Height - 1 : nh);

				k = (double)w * m_h / m_w;
				l = (double)h * m_w / m_h;
				
				l = (l < 0) ? 0 : ((l >= Width) ? Width - 1 : l);
				k = (k < 0) ? 0 : ((k >= Height) ? Height - 1 : k);

				i = (int)k * LineWidth + 3 * (int)l;
				j = (int)nh * LineWidth + 3 * (int)nw;

				NewBits[i+2] = Bits[j+2];
				NewBits[i+1] = Bits[j+1];
				NewBits[ i ] = Bits[ j ];
				Flags[i] = 1;
			}
	
		int m;
		for (w = 1; w < (Width - 1); w++)
			for (int h = 0; h < Height; h++)
			{
				i = h * LineWidth + 3 * w;
				if (! Flags[i])
				{
					j = h * LineWidth + 3 * (w + 1);
					m = h * LineWidth + 3 * (w - 1);
					NewBits[i+2] = (NewBits[j+2] + NewBits[m+2]) / 2;
					NewBits[i+1] = (NewBits[j+1] + NewBits[m+1]) / 2;
					NewBits[ i ] = (NewBits[ j ] + NewBits[ m ]) / 2;
				}
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, NewBits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;
		delete [] NewBits;
		delete [] Flags;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the OilPaint effect (based on Jason Waltman code) 				*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * BrushSize		=> Brush size (duh!)											*
 * Smoothness		=> Smooth value													*
 *																					*
 * Theory			=> Using MostFrequentColor function we take the main color in	*
 *					a matrix and simply write at the original position				*
 *																					*/
HRESULT __stdcall GPX_OilPaint (HDC		PicDestDC, 
								HDC		PicSrcDC,
								int		BrushSize,
								int		Smoothness,
								int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	BrushSize = (BrushSize < 1) ? 1 : (BrushSize > 5) ? 5 : BrushSize;
	Smoothness = (Smoothness < 10) ? 10 : (Smoothness > 255) ? 255 : Smoothness;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE    Bits = new BYTE[BitCount];
	LPBYTE NewBits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, color;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				color = MostFrequentColor (Bits, Width, Height, w, h, BrushSize, Smoothness);
				NewBits[i+2] = (color & 0x000000FF);
				NewBits[i+1] = (color & 0x0000FF00) >>  8;
				NewBits[ i ] = (color & 0x00FF0000) >> 16;
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, NewBits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;
		delete [] NewBits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the FrostGlass effect (based on Jason Waltman code) 			*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Frost			=> Frost Value													*
 *																					*
 * Theory			=> Similar to Diffuse effect, but the random byte is defined	*
 *					in a matrix. Diffuse uses a random diagonal byte.				*
 *																					*/
HRESULT __stdcall GPX_FrostGlass (HDC		PicDestDC, 
								  HDC		PicSrcDC,
								  int		Frost,
								  int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	Frost = (Frost < 1) ? 1 : (Frost > 10) ? 10 : Frost;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE    Bits = new BYTE[BitCount];
	LPBYTE NewBits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, color;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				color = RandomColor (Bits, Width, Height, w, h, Frost);
				NewBits[i+2] = (color & 0x000000FF);
				NewBits[i+1] = (color & 0x0000FF00) >>  8;
				NewBits[ i ] = (color & 0x00FF0000) >> 16;
				DoEvents ();
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, NewBits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;
		delete [] NewBits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the NotePaper effect 											*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Sensibility		=> Sensibility value											*
 * Depth			=> Depth value													*
 * Graininess		=> Grain value													*
 * Intensity		=> Intensity value												*
 * Forecolor		=> Forecolor (duh !)											*
 * Backcolor		=> Backcolor (duh !)											*
 *																					*
 * Theory			=> Its a function thats join Stamp, Rock and a grain effect		*
 *					First, it's applied the stamp effect with forecolor and			*
 *					backcolor (instead Black 'n White), after this, a grain is		*
 *					applied and after, is applied a rock effect.					*
 *																					*/
HRESULT __stdcall GPX_NotePaper (HDC		PicDestDC, 
								 HDC		PicSrcDC,
								 int		Sensibility,
								 int		Depth,
								 int		Graininess,
								 int		Intensity,
								 int		Forecolor,
								 int		Backcolor,
								 int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	BYTE ForeR = (Forecolor & 0x000000FF),
				  ForeG = (Forecolor & 0x0000FF00) >> 8,
				  ForeB = (Forecolor & 0x00FF0000) >> 16,

				  BackR = (Backcolor & 0x000000FF),
				  BackG = (Backcolor & 0x0000FF00) >> 8,
				  BackB = (Backcolor & 0x00FF0000) >> 16;

	if (((ForeR + ForeG + ForeB) / 3) <= 10)
	{
		ForeR += 127;
		ForeG += 127;
		ForeB += 127;
	}

	if (((BackR + BackG + BackB) / 3) <= 10)
	{
		BackR += 127;
		BackG += 127;
		BackB += 127;
	}

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	Sensibility = LimitValues (Sensibility);
	Depth = (Depth < 1) ? 1 : (Depth > 5) ? 5 : Depth;
	Graininess = (Graininess < 0) ? 0 : (Graininess > 60) ? 60 : Graininess;
	Intensity = (Intensity < 0) ? 0 : (Intensity > 10) ? 10 : Intensity;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE    Bits = new BYTE[BitCount];
	LPBYTE NewBits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0, h, w;

		srand ((UINT) GetTickCount());
		for (h = 0; h < Height; h++)
			for (w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				if (((Bits[i+2] + Bits[i+1] + Bits[i]) / 3) < Sensibility)
				{
					Bits[i+2] = ForeR;
					Bits[i+1] = ForeG;
					Bits[ i ] = ForeB;
				}
				else
				{
					Bits[i+2] = BackR;
					Bits[i+1] = BackG;
					Bits[ i ] = BackB;
				}
			}

		int RandValue;
		if (Graininess > 0)
		{
			for (h = 0; h < Height; h++)
				for (w = 0; w < Width; w++)
				{
					i = h * LineWidth + 3 * w;
					if ((Bits[i+2] == BackR) && (Bits[i+1] == BackG) && (Bits[i] == BackB))
					{
						RandValue = (rand() % Graininess);
						if (RandValue % 2)
							RandValue = 0;
						if (((Bits[i+2] + Bits[i+1] + Bits[i]) / 3) > 127)
							RandValue = -RandValue;
						Bits[i+2] = LimitValues (Bits[i+2] + RandValue);
						Bits[i+1] = LimitValues (Bits[i+1] + RandValue);
						Bits[ i ] = LimitValues (Bits[ i ] + RandValue);
					}
					else if ((Bits[i+2] == ForeR) && (Bits[i+1] == ForeG) && (Bits[i] == ForeB))
					{
						RandValue = (rand() % Graininess); 
						if (RandValue % 2)
							RandValue = 0;
						if (((Bits[i+2] + Bits[i+1] + Bits[i]) / 3) > 127)
							RandValue = -RandValue;
						Bits[i+2] = LimitValues (Bits[i+2] + RandValue);
						Bits[i+1] = LimitValues (Bits[i+1] + RandValue);
						Bits[ i ] = LimitValues (Bits[ i ] + RandValue);
					}
				}
		}

		int Shadow;
		for (h = 0; h < Height; h++)
			for (w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				j = (h + Lim_Max (h, Depth, Height)) * LineWidth + 3 * (w + Lim_Max (w, Depth, Width));
				Shadow  = Intensity * (Bits[j+2] - Bits[i+2]);
				Shadow += Intensity * (Bits[j+1] - Bits[i+1]);
				Shadow += Intensity * (Bits[ j ] - Bits[ i ]);
				Shadow /= 3;
				NewBits[i+2] = LimitValues (Bits[i+2] - Shadow);
				NewBits[i+1] = LimitValues (Bits[i+1] - Shadow);
				NewBits[ i ] = LimitValues (Bits[ i ] - Shadow);
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, NewBits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;
		delete [] NewBits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the FishEyeEx effect (based on Jason Waltman code) 			*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Value			=> FishEye value												*
 *																					*
 * Theory			=> This is a great effect, similar to FishEye and Caricature	*
 *					but using polar theories. It's easier to understand than the	*
 *					other FishEye and Caricature effects.							*
 *																					*/
HRESULT __stdcall GPX_FishEyeEx (HDC		PicDestDC, 
								 HDC		PicSrcDC,
								 double		Value,
								 int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (Value == 0)
		return (S_OK);

	Value *= 0.001;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE    Bits = new BYTE[BitCount];
	LPBYTE NewBits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i, j, RadMax = min (Height, Width) / 2;
		int nw, nh;
		int halfH = Height / 2, halfW = Width / 2;
		double r, a;
		double coeff = RadMax / log (fabs (Value) * RadMax + 1);
		for (int h = -1 * halfH; h < Height - halfH; h++)
			for (int w = -1 * halfW; w < Width - halfW; w++)
			{
				r = sqrt (w * w + h * h);
				a = atan2 (h, w);

				i = (h + halfH) * LineWidth + 3 * (w + halfW);

				if (r <= RadMax)
				{
					if (Value > 0)
						r = (exp (r / coeff) - 1) / Value;
					else
						r = coeff * log (1 + (-1 * Value) * r);
					
					nw = halfW + (int)(r * cos (a));
					nh = halfH + (int)(r * sin (a));

					nw = (nw < 0) ? 0 : ((nw >=  Width) ?  Width - 1 : nw);
					nh = (nh < 0) ? 0 : ((nh >= Height) ? Height - 1 : nh);

					j = nh * LineWidth + 3 * nw;
					NewBits[i+2] = Bits[j+2];
					NewBits[i+1] = Bits[j+1];
					NewBits[ i ] = Bits[ j ];
				}
				else
				{
					NewBits[i+2] = Bits[i+2];
					NewBits[i+1] = Bits[i+1];
					NewBits[ i ] = Bits[ i ];
				}
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, NewBits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;
		delete [] NewBits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the RainDrops effect (based on Jason Waltman code) 			*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * DropSize			=> Raindrop size												*
 * Amount			=> Maximum number of raindrops									*
 * Coeff			=> FishEye coefficient											*
 *																					*
 * Theory			=> This functions does several math's functions and the engine	*
 *					is simple to undestand, but a little hard to implement. A		*
 *					control will indicate if there is or not a raindrop in that		*
 *					area, if not, a fisheye effect with a random size (max=DropSize)*
 *					will be applied, after this, a shadow will be applied too.		*
 *					and after this, a blur function will finish the effect.			*
 *																					*/
HRESULT __stdcall GPX_RainDrops (HDC		PicDestDC, 
								 HDC		PicSrcDC, 
								 int		DropSize,
								 int		Amount,
								 int		Coeff,
								 int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (Coeff <= 0)
		Coeff = 1;
	if (Coeff > 100)
		Coeff = 100;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE    Bits = new BYTE[BitCount];
	LPBYTE NewBits = new BYTE[BitCount];
	ppbool BoolMatrix = CreateBoolArray (Width, Height);

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int		i, j, k, l, m, n;				// loop variables

		int		p, q;							// positions

		int		Bright;							// Bright value for shadows and highlights

		int		x, y;							// center coordinates
		double	r, a;							// polar coordinates
		double	OldRadius;						// Radius before processing

		BOOL	FindAnother = false;			// To search for good coordinates
		int		Counter = 0;					// Counter (duh !)

		int		NewSize;						// Size of current raindrop
		int		halfSize;						// Half of the current raindrop
		int		Radius;							// Maximum radius for raindrop

		double	NewCoeff = (double)Coeff * 0.01;// FishEye Coefficients
		double	s;

		int		BlurRadius;						// Blur Radius
		double	R, G, B;
		int		BlurPixels;

		srand ((UINT) GetTickCount());
		for (i = 0; i < Width; i++)
			for (j = 0; j < Height; j++)
			{
				p = j * LineWidth + 3 * i;
				NewBits[p+2] = Bits[p+2];
				NewBits[p+1] = Bits[p+1];
				NewBits[ p ] = Bits[ p ];
				BoolMatrix[i][j] = false;
			}

		for (int NumBlurs = 0; NumBlurs <= Amount; NumBlurs++)
		{
			NewSize = (int)(rand() * ((double)(DropSize - 5) / RAND_MAX) + 5);
			halfSize = NewSize / 2;
			Radius = halfSize;
			s = Radius / log (NewCoeff * Radius + 1);

			Counter = 0;
			do
			{
				FindAnother = false;
				y = (int)(rand() * ((double)( Width - 1) / RAND_MAX));
				x = (int)(rand() * ((double)(Height - 1) / RAND_MAX));

				if (BoolMatrix[y][x])
					FindAnother = true;
				else
					for (i = x - halfSize; i <= x + halfSize; i++)
						for (j = y - halfSize; j <= y + halfSize; j++)
							if ((i >= 0) && (i < Height) && (j >= 0) && (j < Width))
								if (BoolMatrix[j][i])
									FindAnother = true;

				Counter++;
			} while (FindAnother && (Counter < 10000));

			if (Counter >= 10000)
			{
				NumBlurs = Amount;
				break;
			}

			for (i = -1 * halfSize; i < NewSize - halfSize; i++)
			{
				for (j = -1 * halfSize; j < NewSize - halfSize; j++)
				{
					r = sqrt (i * i + j * j);
					a = atan2 (i, j);

					if (r <= Radius)
					{
						OldRadius = r;
						r = (exp (r / s) - 1) / NewCoeff;

						k = x + (int)(r * sin (a));
						l = y + (int)(r * cos (a));

						m = x + i;
						n = y + j;

						if ((k >= 0) && (k < Height) && (l >= 0) && (l < Width))
						{
							if ((m >= 0) && (m < Height) && (n >= 0) && (n < Width))
							{
								p = k * LineWidth + 3 * l;
								q = m * LineWidth + 3 * n;
								NewBits[q+2] = Bits[p+2];
								NewBits[q+1] = Bits[p+1];
								NewBits[ q ] = Bits[ p ];
								BoolMatrix[n][m] = true;
								Bright = 0;
								
								if (OldRadius >= 0.9 * Radius)
								{
									if ((a <= 0) && (a > -2.25))
										Bright = -80;
									else if ((a <= -2.25) && (a > -2.5))
										Bright = -40;
									else if ((a <= 0.25) && (a > 0))
										Bright = -40;
								}

								else if (OldRadius >= 0.8 * Radius)
								{
									if ((a <= -0.75) && (a > -1.50))
										Bright = -40;
									else if ((a <= 0.10) && (a > -0.75))
										Bright = -30;
									else if ((a <= -1.50) && (a > -2.35))
										Bright = -30;
								}

								else if (OldRadius >= 0.7 * Radius)
								{
									if ((a <= -0.10) && (a > -2.0))
										Bright = -20;
									else if ((a <= 2.50) && (a > 1.90))
										Bright = 60;
								}
								
								else if (OldRadius >= 0.6 * Radius)
								{
									if ((a <= -0.50) && (a > -1.75))
										Bright = -20;
									else if ((a <= 0) && (a > -0.25))
										Bright = 20;
									else if ((a <= -2.0) && (a > -2.25))
										Bright = 20;
								}

								else if (OldRadius >= 0.5 * Radius)
								{
									if ((a <= -0.25) && (a > -0.50))
										Bright = 30;
									else if ((a <= -1.75 ) && (a > -2.0))
										Bright = 30;
								}

								else if (OldRadius >= 0.4 * Radius)
								{
									if ((a <= -0.5) && (a > -1.75))
										Bright = 40;
								}

								else if (OldRadius >= 0.3 * Radius)
								{
									if ((a <= 0) && (a > -2.25))
										Bright = 30;
								}

								else if (OldRadius >= 0.2 * Radius)
								{
									if ((a <= -0.5) && (a > -1.75))
										Bright = 20;
								}

								NewBits[q+2] = LimitValues (NewBits[q+2] + Bright);
								NewBits[q+1] = LimitValues (NewBits[q+1] + Bright);
								NewBits[ q ] = LimitValues (NewBits[ q ] + Bright);
							}
						}
					}
				}
			}

			BlurRadius = NewSize / 25 + 1;
			for (i = -1 * halfSize - BlurRadius; i < NewSize - halfSize + BlurRadius; i++)
			{
				for (j = -1 * halfSize - BlurRadius; j < NewSize - halfSize + BlurRadius; j++)
				{
					r = sqrt (i * i + j * j);
					if (r <= Radius * 1.1)
					{
						R = G = B = 0;
						BlurPixels = 0;
						for (k = -1 * BlurRadius; k < BlurRadius + 1; k++)
							for (l = -1 * BlurRadius; l < BlurRadius + 1; l++)
							{
								m = x + i + k;
								n = y + j + l;
								if ((m >= 0) && (m < Height) && (n >= 0) && (n < Width))
								{
									p = m * LineWidth + 3 * n;
									R += NewBits[p+2];
									G += NewBits[p+1];
									B += NewBits[ p ];
									BlurPixels++;
								}
							}

						m = x + i;
						n = y + j;
						if ((m >= 0) && (m < Height) && (n >= 0) && (n < Width))
						{
							p = m * LineWidth + 3 * n;
							NewBits[p+2] = (BYTE)(R / BlurPixels);
							NewBits[p+1] = (BYTE)(G / BlurPixels);
							NewBits[ p ] = (BYTE)(B / BlurPixels);
						}
					}
				}
			}
		}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, NewBits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;
		delete [] NewBits;

		FreeBoolArray (BoolMatrix, Width);
		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the Cilindrical effect 										*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Value			=> Cilindrical value											*
 *																					*
 * Theory			=> This is a great effect, similar to Spherize (Photoshop).		*
 *					If you understand FishEye, you will understand Cilindrical		*
 *					FishEye apply a logarithm function using a sphere radius,		*
 *					Cilindrical use the same function but in a rectangular			*
 *					enviroment.														*
 *																					*/
HRESULT __stdcall GPX_Cilindrical (HDC		PicDestDC, 
								   HDC		PicSrcDC,
								   double	Value,
								   int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (Value == 0)
		return (S_OK);

	Value *= 0.001;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE    Bits = new BYTE[BitCount];
	LPBYTE NewBits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i, j, RadMax = min (Height, Width) / 2;
		int nw, nh;
		int halfH = Height / 2, halfW = Width / 2;
		double r;
		double coeff = RadMax / log (fabs (Value) * RadMax + 1);
		for (int h = -1 * halfH; h < Height - halfH; h++)
			for (int w = -1 * halfW; w < Width - halfW; w++)
			{
				r = fabs ((double)w);

				i = (h + halfH) * LineWidth + 3 * (w + halfW);

				if (r <= RadMax)
				{
					if (Value > 0)
						r = (exp (r / coeff) - 1) / Value;
					else
						r = coeff * log (1 + (-1 * Value) * r);
					
					if (w >= 0)
						nw = halfW + (int)r;
					else
						nw = halfW - (int)r;
					
					nh = halfH + h;

					nw = (nw < 0) ? 0 : ((nw >=  Width) ?  Width - 1 : nw);
					nh = (nh < 0) ? 0 : ((nh >= Height) ? Height - 1 : nh);

					j = nh * LineWidth + 3 * nw;
					NewBits[i+2] = Bits[j+2];
					NewBits[i+1] = Bits[j+1];
					NewBits[ i ] = Bits[ j ];
				}
				else
				{
					NewBits[i+2] = Bits[i+2];
					NewBits[i+1] = Bits[i+1];
					NewBits[ i ] = Bits[ i ];
				}
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, NewBits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;
		delete [] NewBits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the UnsharpMask effect 										*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Blur				=> Blur radius (simple blur)									*
 * Unsharp			=> Depth value for Unsharp										*
 *																					*
 * Theory			=> This effect ins't hard to understand, in a first step, you	*
 *					blur an image, after this, you subtract from original byte the	*
 *					blur byte.														*
 *																					*/
HRESULT __stdcall GPX_UnsharpMask (HDC		PicDestDC, 
								   HDC		PicSrcDC, 
								   short	Blur, 
								   double	Unsharp,
								   int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (Blur < 1)
		return (S_OK);
	if (Blur > 10)
		Blur = 10;

	if (Unsharp < 0.0f)
		return (S_OK);
	if (Unsharp > 10.0f)
		Unsharp = 10.0f;

	int Orig = Blur; 
	Blur = Blur * 2 + 1;
	int Size = Blur * Blur;

	double UnsharpPlus = Unsharp + 1.0f;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE	   Bits = new BYTE[BitCount];
	LPBYTE BlurBits = new BYTE[BitCount];
	LPBYTE	NewBits = new BYTE[BitCount];
	int SomaR = 0, SomaG = 0, SomaB = 0;

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i, j, h, w;
		for (h = 0; h < Height; h++)
			for (w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				for (int a = -Orig; a <= Orig; a++)
					for (int b = -Orig; b <= Orig; b++)
					{
						j = (h + Lim_Max (h, a, Height)) * LineWidth + 3 * (w + Lim_Max (w, b, Width));
						if ((h + a < 0) || (w + b < 0))
							j = i;
						SomaR += Bits[j+2];
						SomaG += Bits[j+1];
						SomaB += Bits[ j ];
					}

				BlurBits[i+2] = SomaR / Size;
				BlurBits[i+1] = SomaG / Size;
				BlurBits[ i ] = SomaB / Size;
				SomaR = SomaG = SomaB = 0;
			}

		for (h = 0; h < Height; h++)
			for (w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				NewBits[i+2] = LimitValues ((int)(UnsharpPlus * Bits[i+2] - Unsharp * BlurBits[i+2]));
				NewBits[i+1] = LimitValues ((int)(UnsharpPlus * Bits[i+1] - Unsharp * BlurBits[i+1]));
				NewBits[ i ] = LimitValues ((int)(UnsharpPlus * Bits[ i ] - Unsharp * BlurBits[ i ]));
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, NewBits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;
		delete [] BlurBits;
		delete [] NewBits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the Flip effect 												*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Width			=> Source image width											*
 * Height			=> Source image height											*
 * Horizontal		=> Flag to indicate that the image will be flipped in horizontal*
 * Vertical			=> Flag to indicate that the image will be flipped in vertical	*
 *																					*
 * Theory			=> With StretchBlt you'll see how easy is to flip an image		*
 *																					*/
HRESULT __stdcall GPX_Flip (HDC  PicDestDC, 
							HDC  PicSrcDC,
							int  Width,
							int  Height,
							BOOL Horizontal,
							BOOL Vertical,
							int	 *Response)
{
	int Result;
	if (Horizontal)
		Result = StretchBlt (PicDestDC, 0, 0, Width, Height, PicSrcDC, Width - 1, 0, -Width, Height, SRCCOPY); 
	
	if (Vertical)
		Result = StretchBlt (PicDestDC, 0, 0, Width, Height, PicSrcDC, 0, Height - 1, Width, -Height, SRCCOPY);

	return (S_OK);
}

/* Function to do the same as BitBlt 												*
 *																					*
 * DestDC			=> Handle to destination DC										*
 * XDest			=> x-coord of destination upper-left corner						*
 * YDest			=> y-coord of destination upper-left corner						*
 * Width			=> Width of destination rectangle								*
 * Height			=> Height of destination rectangle								*
 * SrcDC			=> Handle to source DC											*
 * XSrc				=> x-coord of source upper-left corner							*
 * YSrc				=> y-coord of source upper-left corner							*
 * RasterOp			=> Raster operation code										*
 *																					*
 * Theory			=> Ok, you're right to think "What he is doing ? This function	*
 *					already exists, he just rename it." But think in operability,	*
 *					only you have to do is reference my dll. If you need a function	*
 *					like this, you already have by my dll.							*
 *																					*/
HRESULT __stdcall GPX_BitBlt (HDC DestDC,
							  int XDest,
							  int YDest,
							  int Width,
							  int Height,
							  HDC SrcDC,
							  int XSrc,
							  int YSrc,
							  int RasterOp,
							  int *Response)
{
	*Response = (int)(BitBlt (DestDC, XDest, YDest, Width, Height, SrcDC, XSrc, YSrc, (DWORD)RasterOp)); 
	return (S_OK);
}

/* Function to stretch the histogram image 											*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Flag				=> Indicates what color to be stretched							*
 * StretchFactor	=> Stretch value (1.0 indicates a range from 0 to 255)			*
 *																					*
 * Theory			=> This function is very interesting and not too difficult to	*
 *					implement. Firstly, we get the image histogram, after this, we	*
 *					find the scale value and we calcule in each pixel using this	*
 *					scale value.													*
 *																					*/
HRESULT __stdcall GPX_StretchHistogram (HDC		PicDestDC, 
										HDC		PicSrcDC,
										int		Flag,
										double	StretchFactor,
										int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (StretchFactor > 2.0)
		StretchFactor = 2.0;
	if (StretchFactor < 0.0)
		StretchFactor = 0.0;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, h, w;

		if (Flag & HST_GRAY)
		{
			int GrayPixel;
			for (h = 0; h < Height; h++)
				for (w = 0; w < Width; w++)
				{
					i = h * LineWidth + 3 * w;
					GrayPixel = (Bits[i+2] + Bits[i+1] + Bits[i]) / 3;
					Bits[i+2] = Bits[i+1] = Bits[i] = GrayPixel;
				}

			Flag = HST_COLOR;
		}

		Histogram Red, Green, Blue;
		int ORR, ORG, ORB;
		int SRR, SRG, SRB;
		double ScaleFactorR, ScaleFactorG, ScaleFactorB;

		if (Flag & HST_RED)
		{
			GetHistogram (Bits, Width, Height, Red.Table, HST_RED);
			FindHistoMinAndMaxValues (&Red);
			ORR = Red.Maximum - Red.Minimum;
			SRR = ORR + Round (StretchFactor * (255 - ORR));
			if (! ORR)
				ScaleFactorR = 1.0;
			else
				ScaleFactorR = (double)SRR / (double)ORR;
		}

		if (Flag & HST_GREEN)
		{
			GetHistogram (Bits, Width, Height, Green.Table, HST_GREEN);
			FindHistoMinAndMaxValues (&Green);
			ORG = Green.Maximum - Green.Minimum;
			SRG = ORG + Round (StretchFactor * (255 - ORG));
			if (! ORG)
				ScaleFactorG = 1.0;
			else
				ScaleFactorG = (double)SRG / (double)ORG;
		}

		if (Flag & HST_BLUE)
		{
			GetHistogram (Bits, Width, Height, Blue.Table, HST_BLUE);
			FindHistoMinAndMaxValues (&Blue);
			ORB = Blue.Maximum - Blue.Minimum;
			SRB = ORB + Round (StretchFactor * (255 - ORB));
			if (! ORB)
				ScaleFactorB = 1.0;
			else
				ScaleFactorB = (double)SRB / (double)ORB;
		}

		for (h = 0; h < Height; h++)
			for (w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				if (Flag & 1)
					Bits[i+2] = LimitValues (Round (ScaleFactorR * (Bits[i+2] - Red.Minimum)));
				if (Flag & 2)
					Bits[i+1] = LimitValues (Round (ScaleFactorG * (Bits[i+1] - Green.Minimum)));
				if (Flag & 4)
					Bits[ i ] = LimitValues (Round (ScaleFactorB * (Bits[ i ] - Blue.Minimum)));
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the AlphaBlend effect											*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC_1		=> Source #1 PictureBox's Device Context						*
 * PicSrcDC_2		=> Source #2 PictureBox's Device Context						*
 * Mode				=> Blend mode													*
 *																					*
 * Theory			=> Similar to AlphaBlend function, but has several differences	*
 *					In AlphaBlend function, you use only Average blend mode, here	*
 *					you can use more than 20 blend modes, but, unfortunally, you	*
 *					dont't have a value to make a percentage of these modes.		*
 *																					*/
HRESULT __stdcall GPX_BlendMode (HDC		PicDestDC,
								 HDC		PicSrcDC_1,
								 HDC		PicSrcDC_2,
								 int		Mode,
								 int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width = 0, Height = 0; 
	int SrcWidth = 0, SrcHeight = 0;
	int DestWidth = 0, DestHeight = 0;

	BITMAPINFO infod, infos;
	HBITMAP	PicSrcHwnd_1 = GetBitmapHandle (PicSrcDC_1, &infos, &SrcWidth, &SrcHeight);
	HBITMAP	PicSrcHwnd_2 = GetBitmapHandle (PicSrcDC_2, &infod, &DestWidth, &DestHeight);

	if (! (PicSrcHwnd_1 && PicSrcHwnd_2))
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	Width = (SrcWidth > DestWidth) ? DestWidth : SrcWidth;
	Height = (SrcHeight > DestHeight) ? DestHeight : SrcHeight;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE DestBits = new BYTE[BitCount];
	LPBYTE  SrcBits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC_2, PicSrcHwnd_2, 0, Height, SrcBits, &infos, DIB_RGB_COLORS);
	Result = GetDIBits (PicSrcDC_1, PicSrcHwnd_1, 0, Height, DestBits, &infod, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				DestBits[i+2] = ApplyBlendMode (DestBits[i+2], SrcBits[i+2], Mode);
				DestBits[i+1] = ApplyBlendMode (DestBits[i+1], SrcBits[i+1], Mode);
				DestBits[ i ] = ApplyBlendMode (DestBits[ i ], SrcBits[ i ], Mode);
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, DestBits, &infod, 0);
		
		::DeleteObject (PicSrcHwnd_1);
		::DeleteObject (PicSrcHwnd_2);
		delete [] DestBits;
		delete [] SrcBits;
		
		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the Swirl effect 												*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Swirl			=> Swirl value													*
 *																					*
 * Theory			=> Like FishEye effect, it's very difficult to explain how can	*
 *					I reach on this effect, study hard for this. It's pure			*
 *					trigonometry. If you have spiral theorems, you will understand	*
 *					better, ok?														*
 *																					*/
HRESULT __stdcall GPX_TwirlEx (HDC		PicDestDC, 
							   HDC		PicSrcDC, 
							   double	TwirlMin,
							   double	TwirlMax,
							   int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	double ShiftMin = TwirlMin * PI / 8;
	double ShiftMax = TwirlMax * PI / 8;

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];
	LPBYTE NewBits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, j = 0, nw, nh;
		int halfH = Height / 2, halfW = Width / 2;
		double Angle, Radius, Ratio;
		double Radmax = sqrt (Width * Width + Height * Height);
		for (int w = 0; w < Width; w++)
			for (int h = 0; h < Height; h++)
			{
				nw = w - halfW;
				nh = h - halfH;

				Angle = atan2 (nh, nw);
				Radius = sqrt (nw * nw + nh * nh);
				Ratio = Radius / Radmax;

				i = h * LineWidth + 3 * w;

				if (Ratio > 1.0)
				{
					NewBits[i+2] = Bits[i+2];
					NewBits[i+1] = Bits[i+1];
					NewBits[ i ] = Bits[ i ];
				}
				else
				{
					Angle += Ratio * ShiftMin + (1 - Ratio) * ShiftMax;
					nw = (int)(halfW + Radius * cos (Angle));
					nh = (int)(halfH + Radius * sin (Angle));
					nw = (nw < 0) ? 0 : ((nw >= Width) ? Width - 1 : nw);
					nh = (nh < 0) ? 0 : ((nh >= Height) ? Height - 1 : nh);
					j = nh * LineWidth + 3 * nw;
					NewBits[i+2] = Bits[j+2];
					NewBits[i+1] = Bits[j+1];
					NewBits[ i ] = Bits[ j ];
				}
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, NewBits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;
		delete [] NewBits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply the GlassBlendMode effect										*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC_1		=> Source #1 PictureBox's Device Context						*
 * PicSrcDC_2		=> Source #2 PictureBox's Device Context						*
 * Depth			=> Depth value													*
 * Direction		=> Direction value												*
 *						7 8 1														*
 *						6 0 2														*
 *						5 4 3														*
 *																					*
 * Theory			=> This is a great effect and very easy to undestand, with		*
 *					Alpha value, we can get the proportional color between two		*
 *					pictures.														*
 *																					*/
HRESULT __stdcall GPX_GlassBlendMode (HDC		PicDestDC,
									  HDC		PicSrcDC_1,
									  HDC		PicSrcDC_2, 
									  double	Depth,
									  int		Direction,
									  int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width = 0, Height = 0; 
	int SrcWidth = 0, SrcHeight = 0;
	int DestWidth = 0, DestHeight = 0;
	int neighX = 0, neighY = 0; 

	BITMAPINFO infod, infos;
	HBITMAP	PicSrcHwnd_1 = GetBitmapHandle (PicSrcDC_1, &infos, &SrcWidth, &SrcHeight);
	HBITMAP	PicSrcHwnd_2 = GetBitmapHandle (PicSrcDC_2, &infod, &DestWidth, &DestHeight);

	if (! (PicSrcHwnd_1 && PicSrcHwnd_2))
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	while (Direction > 8)
		Direction -= 8;
	
	if (Direction <= 0)
		return (S_OK);

	if ((Direction >= 1) && (Direction <= 3))
		neighX = 1;
	else if ((Direction >= 5) && (Direction <= 7))
		neighX = -1;

	if ((Direction == 7) || (Direction == 8) || (Direction == 1))
		neighY = 1;
	else if ((Direction >= 3) && (Direction <= 5))
		neighY = -1;

	Width = (SrcWidth > DestWidth) ? DestWidth : SrcWidth;
	Height = (SrcHeight > DestHeight) ? DestHeight : SrcHeight;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE DestBits = new BYTE[BitCount];
	LPBYTE  SrcBits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC_2, PicSrcHwnd_2, 0, Height, DestBits, &infos, DIB_RGB_COLORS);
	Result = GetDIBits (PicSrcDC_1, PicSrcHwnd_1, 0, Height,  SrcBits, &infod, DIB_RGB_COLORS);
	if (Result)
	{
		int i, j, nh, nw, R, G, B, Gray;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				nh = h + neighY;
				nw = w + neighX;
				nw = (nw < 0) ? 0 : ((nw >= Width) ? Width - 1 : nw);
				nh = (nh < 0) ? 0 : ((nh >= Height) ? Height - 1 : nh);
				
				i = h * LineWidth + 3 * w;
				j = nh * LineWidth + 3 * nw;

				R = abs ((int)((SrcBits[i+2] - SrcBits[j+2]) * Depth + 128));
				G = abs ((int)((SrcBits[i+1] - SrcBits[j+1]) * Depth + 128));
				B = abs ((int)((SrcBits[ i ] - SrcBits[ j ]) * Depth + 128));

				Gray = LimitValues ((R + G + B) / 3);

				DestBits[i+2] = LimitValues ((int)DestBits[i+2] + (Gray - 128));
				DestBits[i+1] = LimitValues ((int)DestBits[i+1] + (Gray - 128));
				DestBits[ i ] = LimitValues ((int)DestBits[ i ] + (Gray - 128));
			}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, DestBits, &infod, 0);
		
		::DeleteObject (PicSrcHwnd_1);
		::DeleteObject (PicSrcHwnd_2);
		delete [] DestBits;
		delete [] SrcBits;
		
		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/* Function to apply a metallic layer over the image 											*
 *																					*
 * PicDestDC		=> Destiny PictureBox's Device Context							*
 * PicSrcDC			=> Source PictureBox's Device Context							*
 * Level			=> Metallic levels												*
 * Shift			=> This parameter will shift the metallic table					*
 * Mode				=> Parameter to set what layer will be applied (gold, ice, etc.)*
 *																					*
 * Theory			=> This is one of the most easy functions to understand, this	*
 *					functions takes a pixel and add or sub value to that pixel.		*
 *					As you increase the value, you make the image more bright		*
 *																					*/
HRESULT __stdcall GPX_Metallic (HDC		PicDestDC, 
								HDC		PicSrcDC, 
								int		Level,
								int		Shift,
								int		Mode,
								int		*Response)
{
	int BitCount	= 0;
	int	Result		= 0;
	int Width		= 0; 
	int Height		= 0;
	BITMAPINFO info;
	HBITMAP	PicSrcHwnd = GetBitmapHandle (PicSrcDC, &info, &Width, &Height);

	if (! PicSrcHwnd)
	{
		*Response = (int)Result;
		return (E_FAIL);
	}

	if (Level % 2)
		Level++;

	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	BitCount = LineWidth * Height;
	LPBYTE Bits = new BYTE[BitCount];

	Result = GetDIBits (PicSrcDC, PicSrcHwnd, 0, Height, Bits, &info, DIB_RGB_COLORS);
	if (Result)
	{
		int i = 0, Gray;
		for (int h = 0; h < Height; h++)
			for (int w = 0; w < Width; w++)
			{
				i = h * LineWidth + 3 * w;
				Gray = (Bits[i+2] + Bits[i+1] + Bits[i]) / 3;
				Bits[i+2] = Bits[i+1] = Bits[i] = Gray;
			}

		ApplyMetallicShiftLayer (Bits, Width, Height, Level, Shift);

		switch (Mode)
		{
			case GRAD_GOLD:
				ApplyGoldLayer (Bits, Width, Height);
				break;
			case GRAD_ICE:
				ApplyIceLayer (Bits, Width, Height);
				break;
			default:
				break;
		}

		::SetDIBitsToDevice (PicDestDC, 0, 0, Width, Height, 0, 0, 0, Height, Bits, &info, 0);
		
		::DeleteObject (PicSrcHwnd);
		delete [] Bits;

		*Response = (int)Result;
		return (S_OK);
	}

	return (E_FAIL);
}

/************************************************************************************
 ********                                                                    ********
 ********                               DLL END                              ********
 ********                                                                    ********
 ************************************************************************************/