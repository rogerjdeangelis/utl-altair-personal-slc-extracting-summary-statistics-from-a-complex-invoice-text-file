# utl-altair-personal-slc-extracting-summary-statistics-from-a-complex-invoice-text-file
Altair personal slc extracting summary statistics from a complex invoice text file
    %let pgm=utl-altair-personal-slc-extracting-summary-statistics-from-a-complex-invoice-text-file;

    %stop_submission;

    Too Long to post full solution here.

    github
    https://github.com/rogerjdeangelis/utl-altair-personal-slc-extracting-summary-statistics-from-a-complex-invoice-text-file

    Altair personal slc extracting summary statistics from a complex invoice text file



    PROBLEM (Large text file with repeating invocies, here is one invoice sample (d:/txt/orderstatus.txt) )
    =======================================================================================================

      EXTRACT
            STATE
            DISCOUNT CODE
            STORE ID
            STORE TOTAL
            SALES

      FROM A LARGE NUMBER OF INVOCES


    SAMPLE FROM TEXT FILE
    =====================

      Identify  lines  that look like

      DATE: 04/21/05                ORDER STATUS REPORT BY
      TIME:  8:53:39

      DIV: 0395
                                 ORDER    SHIP
      NAME      ORDER   P.O.#    DATE     DATE     ITEM#
      ----      -----   -----   ------   ------    -----
      KING SEVEN CORPORATION
      P.O. BOX 3456
      DENVER, CO  45607
       34755   95900
                      11/01/04 11/30/04 9S32052S46 YELLOW

                      11/01/04 11/30/04 9S33042S46 WHITE
                      ..

       DISCOUNT CODE: C    STORE ID: 023  STORE TOTAL-->    46                      5184.00

    OUTPUT
    =====

    --------------+
    | A1|fx |STATE|
    --------------------------------------------------------
    [_] |   A     |    B     |    C    |    D    |   E     |
    --------------------------------------------------------
        |         | DISCOUNT | STORE   | STORE   |         |
     1  | STATE   | CODE     |  ID     | TOTAL   | SALES   |
     -- |---------+-----------+---------+---------+--------+
     2  |  CO     | C         | 23      | 46      | 5184   |
     -- |---------+-----------+---------+---------+--------+
     3  |  MA     | A         | 7       | 202     | 20387  |
     -- |---------+-----------+---------+---------+--------+
     4  |  MA     | B         | 15      | 43      | 4454   |
     -- |---------+-----------+---------+---------+--------+
     5  |  IL     | B         | 2       | 56      | 6176   |
     -- |---------+-----------+---------+---------+--------+
    [ORDERSTATUS]


    NOTE: Parmcards4 is not supported in Altair Pesonal SLC (maybe supported in other version of the SLC)
     3        filename ft15f001 "d:/txt/orderstatus.txt";
    34        parmcards4;
              ^
    ERROR: Expected a statement keyword : found "parmcards4"

    /*                   _
    (_)_ __  _ __  _   _| |_
    | | `_ \| `_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    */

    /*--- store the orerstatus in d:/txt/orderstatus.txt ---*/


    data _null_;
      file "d:/txt/orderstatus.txt";
      input;
      put _infile_;
    cards4;
    DATE: 04/21/05                ORDER STATUS REPORT BY CUSTOMER                           PAGE  1
    TIME:  8:53:39

    DIV: 0395

                               ORDER    SHIP
    NAME      ORDER   P.O.#    DATE     DATE     ITEM#    COLOR   QTY  PRICE   AMOUNT   COMMENTS
    ----      -----   -----   ------   ------    -----    -----   ---  -----   ------   --------
    KING SEVEN CORPORATION
    P.O. BOX 3456
    DENVER, CO  45607
              34755   95900
                             11/01/04 11/30/04 9S32052S46 YELLOW   10  100.00  1000.00  SHIPPED FIVE BUT
              36168   90812
                             11/20/04 11/21/04 9S33042S46 BLACK     4  115.00   460.00  SHIPPED

              DISCOUNT CODE: C    STORE ID: 023  STORE TOTAL-->    46                       5184.00


    MONDALE ASSOCIATES
    137 ALEWIFE AVENUE
    STONEHAM, MA 02180

              28382   33902
                             09/01/04 10/30/04 9S3010A    VIOLET    3  108.00   324.00  SHIPPED
              36147   66734
                             11/18/04 11/20/04 9S33014S10 BLACK     1   96.00    96.00  SHIPPED

              DISCOUNT CODE: A    STORE ID: 007  STORE TOTAL-->  202                        20387.00


    BALLARD CORPORATION
    234 BALLARDVALE STREET
    WILMINGTON, MA  01887
              34564   98342
                             11/30/04 11/30/04 9S31158S11 BLACK     2   54.00   108.00  SHIPPED

                             11/30/04 11/30/04 9S38058A11 BLUE      2  155.00   310.00  BACKORDERED
                                                                                        REORDERED

              DISCOUNT CODE: B    STORE ID: 015  STORE TOTAL-->    43                       4454.00

    MONET INC.
    43567 MONET AVENUE
    CHICAGO, IL  26487

              34613   14998
                             11/01/04 11/30/04 9S31165S10 SILVER    2  112.00   224.00  SHIPPED 1 BUT
                             11/01/04 11/30/04 9S39007S31 BLACK     2  102.00   204.00  CANCELLED

              35378   14997
                             11/01/04 11/30/04 9S31036S10 YELLOW    2   72.00   144.00  SHIPPED 1 AND
                                                                                        SHIPPED 1 TO

              DISCOUNT CODE: B    STORE ID: 002  STORE TOTAL-->    56                       6176.00
    ;;;;
    run;quit;

    /*
     _ __  _ __ ___   ___ ___  ___ ___
    | `_ \| `__/ _ \ / __/ _ \/ __/ __|
    | |_) | | | (_) | (_|  __/\__ \__ \
    | .__/|_|  \___/ \___\___||___/___/
    |_|
    */


    &_init_;

    %utlfkil(d:/xls/orderstatusout.xlsx);

    data xls.orderstatus;
      infile "d:/txt/orderstatus.txt";
      input @',' STATE $4. @; /*--- last @ hold the buffer ---*/
      input
          @'DISCOUNT CODE:' discount_code $4.
          @'STORE ID:' store_id  6.
          @'STORE TOTAL-->' store_total $8. sales 32.;
    run;quit;

    proc print width=min;
    run;quit;


    --------------+
    | A1|fx |STATE|
    --------------------------------------------------------
    [_] |   A     |    B     |    C    |    D    |   E     |
    --------------------------------------------------------
        |         | DISCOUNT | STORE   | STORE   |         |
     1  | STATE   | CODE     |  ID     | TOTAL   | SALES   |
     -- |---------+-----------+---------+---------+--------+
     2  |  CO     | C         | 23      | 46      | 5184   |
     -- |---------+-----------+---------+---------+--------+
     3  |  MA     | A         | 7       | 202     | 20387  |
     -- |---------+-----------+---------+---------+--------+
     4  |  MA     | B         | 15      | 43      | 4454   |
     -- |---------+-----------+---------+---------+--------+
     5  |  IL     | B         | 2       | 56      | 6176   |
     -- |---------+-----------+---------+---------+--------+
    [ORDERSTATUS]

    /*
    | | ___   __ _
    | |/ _ \ / _` |
    | | (_) | (_| |
    |_|\___/ \__, |
             |___/
    */

    220       %utlfkil(d:/xls/orderstatusout.xlsx);
    221
    222       data xls.orderstatus;
    223         infile "d:/txt/orderstatus.txt";
    224         input @',' STATE $4. @; /*--- last @ hold the buffer ---*/
    225         input
    226             @'DISCOUNT CODE:' discount_code $4.
    227             @'STORE ID:' store_id  6.
    228             @'STORE TOTAL-->' store_total $8. sales 32.;
    229       run;

    NOTE: The infile 'd:\txt\orderstatus.txt' is:
          Filename='d:\txt\orderstatus.txt',
          Owner Name=T7610\Roger,
          File size (bytes)=6480,
          Create Time=13:13:09 Oct 18 2025,
          Last Accessed=13:30:24 Oct 18 2025,
          Last Modified=13:13:57 Oct 18 2025,
          Lrecl=32767, Recfm=V

    NOTE: 40 records were read from file 'd:\txt\orderstatus.txt'
          The minimum record length was 160
          The maximum record length was 160
    NOTE: Data set "XLS.orderstatus" has an unknown number of observation(s) and 5 variable(s)
    NOTE: The data step took :
          real time : 0.711
          cpu time  : 0.531


    229     !     quit;
    230
    231       proc print width=min;
    232       run;quit;
    NOTE: 4 observations were read from "XLS.orderstatus"
    NOTE: Procedure print step took :
          real time : 0.352
          cpu time  : 0.343

    /*              _
      ___ _ __   __| |
     / _ \ `_ \ / _` |
    |  __/ | | | (_| |
     \___|_| |_|\__,_|

    */
