#include "GlobalFunctions.h"

/************************************************************************************
 ********                                                                    ********
 ********                      NOT EXPORTED FUNCTIONS                        ********
 ********                                                                    ********
 ************************************************************************************/

/* This function does the same as DoEvents in VB enviroment							*
 *																					*
 * No parameters																	*/
void DoEvents (void) 
{ 
    MSG Msg; 
    while (PeekMessage (&Msg, NULL, 0, 0, PM_REMOVE))
    {
        if (Msg.message == WM_QUIT)
			break;
        TranslateMessage (&Msg); 
        DispatchMessage (&Msg);
    } 
}

/* This function does the same as Round() in VB enviroment							*
 *																					*
 * No parameters																	*/
int Round (double Value) 
{ 
    return ((int)floor (Value + 0.5));
}

/* This function limits the RGB values												*
 *																					*
 * ColorValue		=> Here, is an RGB value to be analized							*
 *																					*
 * Theory			=> A color is represented in RGB value (e.g. 0xFFFFFF is		*
 *					white color). But R, G and B values has 256 values to be used	*
 *					so, this function analize the value and limits to this range	*
 *																					*/					   
inline BYTE LimitValues (int ColorValue)
{
	if (ColorValue > 255)		// MAX = 255
		ColorValue = 255;		
	if (ColorValue < 0)			// MIN = 0
		ColorValue = 0;
	return ((BYTE) ColorValue);
}

/* This function limits the max and min values 										*
 * defined by the developer															*
 *																					*
 * Now				=> Original value												*
 * Up				=> Increments													*
 * Max				=> Maximum value												*
 *																					*
 * Theory			=> This function is used in some functions to limit the			*
 *					"for step". E.g. I have a picture with 309 pixels (width), and	*
 *					my "for step" is 5. All the code go alright until reachs the	*
 *					w = 305, because in the next step w will go to 310, but we want	*
 *					to analize all the pixels. So, this function will reduce the	*
 *					"for step", when necessary, until reach the last possible value	*
 *																					*/
inline int Lim_Max (int Now, 
					int Up, 
					int Max)
{
	Max--;
	while (Now > Max - Up)
		Up--;
	return (Up);
}

/* ShadeColors function (based on Martin code)										*
 *																					*
 * DestColor		=> Destiny colour												*
 * SrcColor			=> Source colour												*
 * Shade			=> Value to find the intermediate								*
 *																					*
 * Theory			=> This is a great code thats calcule the intermediate value	*
 *					Think thats 255 (shade parameter) is 100% of SourceColor		*
 *																					*/
inline int ShadeColors (int DestColor, 
						int SrcColor, 
						int Shade)
{
	if (Shade == 0)			// if shade is 0 return the DestColor
		return DestColor;
	if (Shade == 255)		// if shade is 255 return the SrcColor
		return SrcColor;
	// Now, calcule the intermediate colour
	return ((DestColor * (255 - Shade) + SrcColor * Shade) >> 8);
}

/* ProportionalValue function														*
 *																					*
 * DestColor		=> Destiny colour												*
 * SrcColor			=> Source colour												*
 * Shade			=> Value to find the intermediate								*
 *																					*
 * Theory			=> This function does the same thing that ShadeColors function	*
 *					but using double variables.										*
 *																					*/
inline double ProportionalValue (double DestValue, 
								 double SrcValue, 
								 double Shade)
{
	if (Shade == 0.0)			// if shade is 0 return the DestColor
		return DestValue;
	if (Shade == 255.0)		// if shade is 255 return the SrcColor
		return SrcValue;
	// Now, calcule the intermediate colour
	return ((DestValue * (255.0 - Shade) + SrcValue * Shade) / 256.0);
}

/* Function to calcule the color intensity											*
 *																					*
 * R				=> Red value													*
 * G				=> Green value													*
 * B				=> Blue value													*
 *																					*
 * Theory			=> This is the luminance (Y) component of YIQ color model		*
 *																					*/
inline BYTE GetIntensity (BYTE R,
						  BYTE G,
						  BYTE B)
{
	return ((BYTE)(R * 0.3 + G * 0.59 + B * 0.11));
}

/* Function to return the maximum radius with a determined angle					*
 *																					*
 * Height			=> Height of the image											*
 * Width			=> Width of the image											*
 * Angle			=> Angle to analize the maximum radius							*
 *																					*
 * Theory			=> This function calcule the maximum radius to that angle		*
 *					so, we can build an oval circunference							*
 *																					*/
inline double MaximumRadius (int Height, 
							 int Width, 
							 double Angle)
{
	double MaxRad, MinRad;
	double Radius, DegAngle = fabs (Angle * 57.295);	// Rads -> Degrees

	MinRad = min (Height, Width) / 2.0;					// Gets the minor radius
	MaxRad = max (Height, Width) / 2.0;					// Gets the major radius

	// Find the quadrant between -PI/2 and PI/2
	if (DegAngle > 90.0)
		Radius = ProportionalValue (MinRad, MaxRad, (DegAngle * (255.0 / 90.0)));
	else
		Radius = ProportionalValue (MaxRad, MinRad, ((DegAngle - 90.0) * (255.0 / 90.0)));
	return (Radius);
}

/* Function to determine the most frequent color in a matrix						*
 *																					*
 * *Bits			=> Bits array													*
 * Width			=> Image width													*
 * Height			=> Image height													*
 * X				=> Position horizontal											*
 * Y				=> Position vertical											*
 * Radius			=> Is the radius of the matrix to be analized					*
 * Intensity		=> Intensity to calcule											*
 *																					*
 * Theory			=> This function creates a matrix with the analized pixel in	*
 *					the center of this matrix and find the most frequenty color		*
 *																					*/
inline int MostFrequentColor (LPBYTE	Bits, 
						      int		Width, 
						      int		Height, 
						      int		X, 
						      int		Y, 
						      int		Radius, 
						      int		Intensity)
{
	int i, w, h, color;
	double Scale = Intensity / 255.0;
	int LineWidth = 3 * Width;
	if (LineWidth % 4)						// Don't take off this step
		LineWidth += (4 - LineWidth % 4);
	BYTE I;

	// Alloc some arrays to be used
	LPBYTE	IntensityCount	= new BYTE[(Intensity + 1) * sizeof (BYTE)];
	UINT	*AverageColorR  = new UINT[(Intensity + 1) * sizeof (UINT)];
	UINT	*AverageColorG  = new UINT[(Intensity + 1) * sizeof (UINT)];
	UINT	*AverageColorB  = new UINT[(Intensity + 1) * sizeof (UINT)];

	// Erase the array
	for (i = 0; i <= Intensity; i++)
		IntensityCount[i] = 0;

	for (w = X - Radius; w <= X + Radius; w++)
		for (h = Y - Radius; h <= Y + Radius; h++)
		{
			// This condition helps to identify when a point doesn't exist
			if ((w >= 0) && (w < Width) && (h >= 0) && (h < Height))
			{
				// You'll see a lot of times this formula
				i = h * LineWidth + 3 * w;
				I = (BYTE)(GetIntensity (Bits[i+2], Bits[i+1], Bits[i]) * Scale);
				IntensityCount[I]++;

				if (IntensityCount[I] == 1)
				{
					AverageColorR[I] = Bits[i+2];
					AverageColorG[I] = Bits[i+1];
					AverageColorB[I] = Bits[ i ];
				}
				else
				{
					AverageColorR[I] += Bits[i+2];
					AverageColorG[I] += Bits[i+1];
					AverageColorB[I] += Bits[ i ];
				}
			}
		}

	I = 0;
	int MaxInstance = 0;

	for (i = 0; i <= Intensity; i++)
		if (IntensityCount[i] > MaxInstance)
		{
			I = i;
			MaxInstance = IntensityCount[i];
		}

	int R, G, B;
	R = AverageColorR[I] / MaxInstance;
	G = AverageColorG[I] / MaxInstance;
	B = AverageColorB[I] / MaxInstance;
	color = RGB (R, G, B);

	delete [] IntensityCount;		// free all the arrays
	delete [] AverageColorR;
	delete [] AverageColorG;
	delete [] AverageColorB;

	return (color);					// return the most frequenty color
}

/* Function to get a color in a matriz with a determined size						*
 *																					*
 * *Bits			=> Bits array													*
 * Width			=> Image width													*
 * Height			=> Image height													*
 * X				=> Position horizontal											*
 * Y				=> Position vertical											*
 * Radius			=> The radius of the matrix to be created						*
 *																					*
 * Theory			=> This function takes from a distinct matrix a random color	*
 *																					*/
inline int RandomColor (LPBYTE	Bits, 
						int		Width, 
						int		Height, 
						int		X, 
						int		Y, 
						int		Radius)
{
	int i, w, h, color, counter = 0;
	int LineWidth = 3 * Width;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);
	BYTE I;
	const BYTE MAXINTENSITY = 255;

	BYTE IntensityCount[MAXINTENSITY + 1];
	UINT AverageColorR[MAXINTENSITY + 1];
	UINT AverageColorG[MAXINTENSITY + 1];
	UINT AverageColorB[MAXINTENSITY + 1];

	for (i = 0; i <= MAXINTENSITY; i++)
		IntensityCount[i] = 0;

	for (w = X - Radius; w <= X + Radius; w++)
		for (h = Y - Radius; h <= Y + Radius; h++)
		{
			if ((w >= 0) && (w < Width) && (h >= 0) && (h < Height))
			{
				i = h * LineWidth + 3 * w;
				I = (BYTE)(GetIntensity (Bits[i+2], Bits[i+1], Bits[i]));
				IntensityCount[I]++;
				counter++;

				if (IntensityCount[I] == 1)
				{
					AverageColorR[I] = Bits[i+2];
					AverageColorG[I] = Bits[i+1];
					AverageColorB[I] = Bits[ i ];
				}
				else
				{
					AverageColorR[I] += Bits[i+2];
					AverageColorG[I] += Bits[i+1];
					AverageColorB[I] += Bits[ i ];
				}
			}
		}

	int RandNumber, count, Index, ErrorCount = 0;

	do
	{
		RandNumber = (int)((rand() + 1) * ((double)counter / (RAND_MAX + 1)));
		count = 0;
		Index = 0;
		do
		{
			count += IntensityCount[Index];
			Index++;
		} while (count < RandNumber);

		I = Index - 1;
		ErrorCount++;

	} while ((IntensityCount[I] == 0) && (ErrorCount <= counter));

	int R, G, B;

	if (ErrorCount >= counter)
	{
		R = AverageColorR[I] / counter;
		G = AverageColorG[I] / counter;
		B = AverageColorB[I] / counter;
	}
	else
	{
		R = AverageColorR[I] / IntensityCount[I];
		G = AverageColorG[I] / IntensityCount[I];
		B = AverageColorB[I] / IntensityCount[I];
	}
	color = RGB (R, G, B);

	return (color);
}

/* Function to free a dinamic boolean array											*
 *																					*
 * lpbArray			=> Dinamic boolean array										*
 * Columns			=> The array bidimension value									*
 *																					*
 * Theory			=> An easy to undestand 'for' statement							*
 *																					*/
void FreeBoolArray (ppbool lpbArray, UINT Columns)
{
	for (UINT i = 0; i < Columns; i++)
		free (lpbArray[i]);
	free (lpbArray);
}

/* Function to create a bidimentional dinamic boolean array							*
 *																					*
 * Columns			=> Number of columns											*
 * Rows				=> Number of rows												*
 *																					*
 * Theory			=> Using 'for' statement, we can alloc multiple dinamic arrays	*
 *					To create more dimentions, just add some 'for's, ok?			*
 *																					*/
ppbool CreateBoolArray (UINT Columns, UINT Rows)
{
	ppbool lpbArray = NULL;
	lpbArray = (ppbool) malloc (Columns * sizeof (lpbool));

	if (lpbArray == NULL)
		return (NULL);

	for (UINT i = 0; i < Columns; i++)
	{
		lpbArray[i] = (lpbool) malloc (Rows * sizeof (bool));
		if (lpbArray[i] == NULL)
		{
			FreeBoolArray (lpbArray, Columns);
			return (NULL);
		}
	}

	return (lpbArray);
}

/* Function to get a HBITMAP from a Device Context, and returning the dimentions	*
 *																					*
 * hDC				=> Device Context to be analized								*
 * bmpInfo			=> Bitmap info to be returned									*
 * Width			=> Return the bitmap's width									*
 * Height			=> Return the bitmap's height									*
 *																					*
 * Theory			=> Here, we get a bitmap object from a device context, after	*
 *					this, we initialize the bitmap info with the needed values		*
 *																					*/
HBITMAP GetBitmapHandle (HDC hDC, BITMAPINFO *bmpInfo, int *Width, int *Height)
{
	int biSize = sizeof (BITMAPINFOHEADER);
	HBITMAP hBitmap; 
	BITMAP bmp;
	
	ZeroMemory (bmpInfo, biSize);
	bmpInfo->bmiHeader.biSize = biSize;

	hBitmap = (HBITMAP)::GetCurrentObject (hDC, OBJ_BITMAP);
	::GetObject (hBitmap, sizeof (BITMAP), (LPSTR)&bmp);

	if (! ::GetDIBits (hDC, hBitmap, 0, 0, NULL, bmpInfo, DIB_RGB_COLORS))
		return (NULL);
	
	*Width	= bmpInfo->bmiHeader.biWidth,
	*Height	= bmpInfo->bmiHeader.biHeight;
	
	bmpInfo->bmiHeader.biCompression	= 0;
	bmpInfo->bmiHeader.biBitCount		= 24;

	return (hBitmap);
}

/* Function to assign table colors to a bitmap										*
 *																					*
 * RedTable			=> Table for red color											*
 * GreenTable		=> Table for green color										*
 * BlueTable		=> Table for blue color											*
 * Bits				=> bitmap bits													*
 * Width			=> Bitmap's width												*
 * Height			=> Bitmap's height												*
 *																					*
 * Theory			=> We change the colors with a table association				*
 *																					*/
void AssignTables (UCHAR	*RedTable, 
				   UCHAR	*GreenTable, 
				   UCHAR	*BlueTable, 
				   UCHAR	Bits[],  
				   int		Width, 
				   int		Height)
{
	int LineWidth = Width * 3;
	int Stride = (LineWidth % 4) ? 4 - LineWidth % 4 : 0;
	int h, w, i = 0;

	for (h = 0; h < Height; h++, i += Stride)
		for (w = 0; w < Width; w++, i += 3)
		{
			Bits[i+2] =   RedTable[Bits[i+2]];
			Bits[i+1] = GreenTable[Bits[i+1]];
			Bits[ i ] =  BlueTable[Bits[ i ]];
		}

	return;
}

/* Function not yet used (for future use)											*
 *																					*
 *																					*/
UCHAR *MakeConvolution (UCHAR Bits[],
					    int Width,
					    int Height,
					    double Coeffs[],
					    int Dimention)
{
	int Size = Dimention * Dimention;
	int LineWidth = Width * 3;
	if (LineWidth % 4)
		LineWidth += (4 - LineWidth % 4);

	if (Dimention < 1)
		return (Bits);

	UCHAR *Temp = new UCHAR[LineWidth * Height];

	int i, j, pos;
	double sumR = 0, sumG = 0, sumB = 0;
	for (int h = Dimention; h < Height - Dimention; h++)
		for (int w = Dimention; w < Width - Dimention; w++)
		{
			i = h * LineWidth + 3 * w;
			for (int a = -Dimention; a <= Dimention; a++)
				for (int b = -Dimention; b <= Dimention; b++)
				{
					j = (h + a) * LineWidth + 3 * (w + b);
					pos = ((a + Dimention) * (Dimention * 2 + 1) + (b + Dimention));
					sumR += (Bits[j+2] * Coeffs[pos]);
					sumG += (Bits[j+1] * Coeffs[pos]);
					sumB += (Bits[ j ] * Coeffs[pos]);
				}

			Temp[i+2] = (UCHAR)LimitValues ((int)sumR);
			Temp[i+1] = (UCHAR)LimitValues ((int)sumG);
			Temp[ i ] = (UCHAR)LimitValues ((int)sumB);
			sumR = sumG = sumB = 0;
		}

	return (Temp);
}

