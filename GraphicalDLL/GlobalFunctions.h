#include <windows.h>
#include <math.h>
#include <stdlib.h>
#include <tchar.h>
#include <malloc.h>

#define COLOR_SIZE			256

/************************************************************************************
 ********                                                                    ********
 ********                               TYPEDEFS                             ********
 ********                                                                    ********
 ************************************************************************************/

typedef char*	string;				// Definition of string (for some functions here, ok?)
typedef bool*	lpbool;
typedef bool**	ppbool;

/************************************************************************************
 ********                                                                    ********
 ********                              FUNCTIONS                             ********
 ********                                                                    ********
 ************************************************************************************/

void	DoEvents	(void);
int		Round		(double Value);

extern BYTE		LimitValues			(int ColorValue);
extern int		Lim_Max				(int Now, int Up, int Max);
extern int		ShadeColors			(int DestColor, int SrcColor, int Shade);
extern double	ProportionalValue	(double DestValue, double SrcValue, double Shade);
extern BYTE		GetIntensity		(BYTE R, BYTE G, BYTE B);
extern double	MaximumRadius		(int Height, int Width, double Angle);
extern int		MostFrequentColor	(LPBYTE	Bits, int Width, int Height, 
									 int X, int Y, int Radius, int Intensity);
extern int		RandomColor			(LPBYTE	Bits, int Width, int Height, 
									 int X, int Y, int Radius);

void	FreeBoolArray		(ppbool lpbArray, UINT Columns);
ppbool	CreateBoolArray		(UINT Columns, UINT Rows);
HBITMAP GetBitmapHandle		(HDC hDC, BITMAPINFO *bmpInfo, int *Width, int *Height);
void	AssignTables		(UCHAR *RedTable, UCHAR	*GreenTable, UCHAR *BlueTable, 
							 UCHAR Bits[], int Width, int Height);
UCHAR*	MakeConvolution		(UCHAR Bits[], int Width, int Height,
							 double *Coeffs, int Dimention);
