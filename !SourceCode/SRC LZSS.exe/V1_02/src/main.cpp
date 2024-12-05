// Test exe file for compression classes - very untidy code only used
// for testing compression functions quickly
//
// CONSOLE APP
//
// (c)2002-2003 Jonathan Bennett, jon@hiddensoft.com
//


#include <stdio.h>
#include <conio.h>
#include <windows.h>
#include <mmsystem.h>
#include "jb01_compress.h"
#include "jb01_decompress.h"

static	UINT  EXITCODE_OK = 0;
static	UINT  EXITCODE_ERR = 1;
static	UINT  EXITCODE_ERR_INPUTPARAMS = 2;
static	UINT  EXITCODE_ERR_EXCEP = 3;



///////////////////////////////////////////////////////////////////////////////
// CompressMonitorProc() - The callback function
///////////////////////////////////////////////////////////////////////////////

int CompressMonitorProc(ULONG nBytesIn, ULONG nBytesOut, UINT nPercentComplete)
{
	static	UINT	nDelay = 0;
	static	UINT	nRot = 0;
	char	szGfx[]= "-\\|/";
	UCHAR	ch;

//	if (nDelay > 16)
//	{
		nDelay = 0;
		nRot = (nRot+1) & 0x3;
		printf("\rCompressing %c        : %d%% (%d%%)  ", szGfx[nRot], nPercentComplete, 100-((100*nBytesOut) / nBytesIn));

//	}
//	else
//		nDelay++;

	// Check if ESC was pressed and if so request stopping
	if (_kbhit())
	{
		ch = _getch();
		if (ch == 0)
			ch = _getch();

		if (ch == 27)
			return 0;
	}

	return 1;

} // CompressProc()


///////////////////////////////////////////////////////////////////////////////
// DeompressMonitorProc() - The callback function
///////////////////////////////////////////////////////////////////////////////

int DecompressMonitorProc(ULONG nBytesIn, ULONG nBytesOut, UINT nPercentComplete)
{
	static	UINT	nDelay = 0;
	static	UINT	nRot = 0;
	char	szGfx[]= "-\\|/";
	UCHAR	ch;

//	if (nDelay > 16)
//	{
		nDelay = 0;
		nRot = (nRot+1) & 0x3;
		printf("\rDecompressing %c      : %d%%  ", szGfx[nRot], nPercentComplete);
//	}
//	else
//		nDelay++;

	// Check if ESC was pressed and if so request stopping
	if (_kbhit())
	{
		ch = _getch();
		if (ch == 0)
			ch = _getch();

		if (ch == 27)
			return 0;
	}

	return 1;

} // DecompressMonitorProc()


///////////////////////////////////////////////////////////////////////////////
// main()
///////////////////////////////////////////////////////////////////////////////

int main(int argc, char* argv[])
{

  //try {
	  unsigned long		nCompressedSize;
	  unsigned long		nUncompressedSize;
	  int					nRes;
	  JB01_Compress	oCompress;					// Our compression class
	  JB01_Decompress	oDecompress;				// Our decompression class
    const long NUM_ARGS_REQUIRED = 4;

	  printf("\nHiddenSoft Compression Routine - (c)2002-2003 Jonathan Bennett\n");
	  printf("--> Extended Version 0.3 by CW2K - [16.11.2016]\n");
	  printf("--------------------------------------------------------------\n\n");

	  // Compress file to file function
    if (argc == NUM_ARGS_REQUIRED) {

      if ( !_stricmp("-c", argv[1]) )
      {
        // How big is the source file?
        nUncompressedSize = oCompress.GetFileSize(argv[2]);
        printf("Input file size      : %d\n", nUncompressedSize);

        // Do the compression
        oCompress.SetDefaults();
        oCompress.SetInputType(   HS_COMP_FILE);
        oCompress.SetOutputType   (HS_COMP_FILE);
        oCompress.SetInputFile(   argv[2]);
        oCompress.SetOutputFile(  argv[3]);
        oCompress.SetMonitorCallback( &CompressMonitorProc);

        oCompress.SetCompressionLevel(4);


        DWORD dwTime1 = timeGetTime();

        nRes = oCompress.Compress();

        DWORD dwTime2 = timeGetTime();


        printf("\rCompressed           : %d%% (%d%%)  ",  oCompress.GetPercentComplete(),
          100 - (oCompress.GetCompressedSize() / nUncompressedSize) * 100);

        printf("\nCompression time     : %.2fs (including fileIO)\n",
          ((dwTime2 - dwTime1)) / 1000.0);

        if (nRes != JB01_E_OK)
        {
          printf("Error compressing.\n");
          return nRes;
        }

        // Print the output size
        printf("Output file size     : %d\n",              oCompress.GetCompressedSize() );
        printf("Compression ratio    : %.2f%%\n",  100 - ((oCompress.GetCompressedSize()  / nUncompressedSize) * 100));
        printf("Compression ratio    : %.3f bpb\n", (8 *   oCompress.GetCompressedSize()) / nUncompressedSize);

        return EXITCODE_OK;
      }




      // Uncompress file to file function
      if ( !_stricmp("-d", argv[1]) )
      {
        // How big is the source file?
        nCompressedSize = oCompress.GetFileSize(argv[2]);
        printf("Input file size      : %8d\n", nCompressedSize);
        printf("Size decompressed    : %8d\n", oDecompress.GetDecompressedSize() );

        // Do the uncompression
        oDecompress.SetDefaults();
        oDecompress.SetInputType(   HS_COMP_FILE  );
        oDecompress.SetOutputType(  HS_COMP_FILE  );
        oDecompress.SetInputFile(     argv[2] );
        oDecompress.SetOutputFile(    argv[3] );
        oDecompress.SetMonitorCallback( &DecompressMonitorProc);


        DWORD dwTime1 = timeGetTime();

        nRes = oDecompress.Decompress();

        DWORD dwTime2 = timeGetTime();

        printf("\rDecompressed         : %8d%%  ", oDecompress.GetPercentComplete());
        printf("\nCompression time     : %8.2fs (including fileIO)\n", ((dwTime2 - dwTime1)) / 1000.0);

        if (nRes != JB01_E_OK)
        {
          printf("Error uncompressing.\n");

          switch (nRes) {
            case JB01_E_MEMALLOC:            printf("Memory alloc failed.\n"); break;
            case JB01_E_READINGSRC:          printf("Read error on inputfile.\n"); break;
            case JB01_E_READINGSRCTRUNCATED: printf(
              "     Inputfile got truncated.\n"
              "     Only %8d of %8d bytes were decompressed."
              , oDecompress.GetDecompressedDataWritten(),
              oDecompress.GetDecompressedSize()
              ); break;
          }


          return nRes;
        }

        // Print filesize
        printf("Output file size     : %d\n", oCompress.GetFileSize(argv[3]));

        return EXITCODE_OK;

      }
    }
    else
      printf("Notice: No action performed - got %d of %d of the required commandline arguments! \n",
        argc, NUM_ARGS_REQUIRED);



	  // If we got here, invalid parameters
	  printf("Usage: %s <-c | -d> <infile> <outfile>\n", argv[0]);
	  printf("  -c performs file to file compression\n");
	  printf("  -d performs file to file decompression\n\n");
	  printf("Supported files type(s) 'JB01' (and 'JB00', 'EA05', 'EA06' decompression only).\n");
 // }
 // catch (EXCEPINFO e) {
 //   return EXITCODE_ERR_EXCEP;
 // }

	return EXITCODE_ERR_INPUTPARAMS;
}
