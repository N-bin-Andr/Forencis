package Bones.Radius;

import Bones.Bones;

public class Radius extends Bones {



    //ARRAYS:
    //MANOUVRIER
    private static  final float[]MANOUVRIER_mLenght = {
            //длина лучевой кости мужчин (рост 153-183см).
            213f, 216f, 219f, 222f, 225f, 229f, 232f, 236f, 239f, 243f, 246f,
            249f, 252f, 255f, 258f, 261f, 264f, 267f, 270f, 273f
    };
    private static final float[] MANOUVRIER_wLenght = {
            //длина лучевой кости женщин (рост 140-171,5см).
            193f, 195f, 197f, 199f, 201f, 203f, 205f, 207f, 209f, 211f, 214f,
            218f, 222f, 226f, 230f, 234f, 238f, 242f, 246f, 250f
    };

    //TROTTER_&_GLESER
    private float[]mGrPr_TrotterGleser;//рост мужчины.
    private float[]wGrPr_TrotterGleser;//рост женщины.
    private static final float [] mL_TROTTER_GLESER={
            //длина лучевой  кости мужчин (рост 152-198см.):
            193f, 196f, 198f, 201f, 204f, 206f, 209f, 212f, 214f, 217f, 220f,
            222f, 225f, 228f, 230f, 233f, 235f, 238f, 241f, 243f, 246f, 249f,
            251f, 254f, 257f, 259f, 262f, 265f, 267f, 270f, 272f, 275f, 278f,
            280f, 283f, 286f, 288f, 291f, 294f, 296f, 299f, 302f, 304f, 307f,
            309f, 312f, 315f

/* по Алексееву (1966) на основании TrotterGleser
    длина лучевой кости мужчин (европеоиды) (рост 152-198см.):
            192f, 194f, 197f, 199f, 202f, 205f, 207f, 210f, 213f, 215f, 218f, 221f,
            223f, 226f, 228f, 231f, 234f, 236f, 239f, 242f, 244f, 247f, 250f, 252f,
            255f, 257f, 260f, 263f, 265f, 268f, 271f, 273f, 276f, 279f, 281f, 284f,
            286f, 289f, 292f, 294f, 297f, 300f, 302f, 305f, 308f, 310f, 313f

     длина лучевой кости мужчин (монголоиды) (рост 152-198см.):
            198f, 201f, 203f, 206f, 209f, 212f, 215f, 218f, 220f, 223f, 226f, 229f,
            232f, 234f, 237f, 240f, 243f, 246f, 249f, 251f, 254f, 257f, 260f, 263f,
            266f, 268f, 271f, 274f, 277f, 280f, 282f, 285f, 288f, 291f, 294f, 297f,
            299f, 302f, 305f, 308f, 311f, 314f, 316f, 319f, 322f, 325f, 328f

     длина лучевой кости мужчин (негроиды) (рост 152-198см.):
            201f, 204f, 207f, 210f, 213f, 216f, 219f, 222f, 225f, 228f, 231f, 234f,
            237f, 240f, 243f, 246f, 249f, 252f, 255f, 258f, 261f, 264f, 267f, 270f,
            273f, 276f, 279f, 282f, 285f, 288f, 291f, 294f, 297f, 300f, 303f, 306f,
            309f, 312f, 315f, 318f, 321f, 324f, 327f, 331f, 334f, 337f, 340f
 */

    } ;
    private static final float[]wL_TROTTER_GLESER={
            //длина лучевой кости женжин (европеоиды) (рост 140-184см.):
            179f, 182f, 184f, 186f, 188f, 190f, 192f, 194f, 196f, 198f, 201f,
            203f, 205f, 207f, 209f, 211f, 213f, 215f, 217f, 220f, 222f, 224f,
            226f, 228f, 230f, 232f, 234f, 236f, 239f, 241f, 243f, 245f, 247f,
            249f, 251f, 253f, 255f, 258f, 260f, 262f, 264f, 266f, 268f, 270f,
            272f
    };
    private static final float[]wnL_TROTTER_GLESER={
            //длина лучевой кости женщин  (негроиды) (рост 140-184см.):
            165f, 169f, 173f, 176f, 180f, 184f, 187f, 191f, 195f, 198f, 202f,
            205f, 209f, 213f, 216f, 220f, 224f, 227f, 231f, 235f, 238f, 242f,
            245f, 249f, 253f, 256f, 260f, 264f, 267f, 271f, 275f, 278f, 282f,
            285f, 289f, 293f, 296f, 300f, 304f, 307f, 311f, 315f, 318f, 322f,
            325f
    };

    //TELCCA
    private float[]mGrPr_Telcca;//рост мужчины.
    private float[]wGrPr_Telcca;//рост женщины.
    private static final  float [] mL_TELCCA={
    //длина лучевой кости  мужчин (рост 152-198см.):
            185f, 188f, 191f, 194f, 197f, 199f, 202f, 205f, 208f, 211f, 214f,
            217f, 220f, 223f, 226f, 229f, 232f, 235f, 238f, 241f, 244f, 246f,
            249f, 252f, 255f, 258f, 261f, 264f, 267f, 270f, 273f
    };
    private static final float [] wL_TELCCA = {
    //длина женской кости.
            170f, 173f, 176f, 180f, 183f, 186f, 189f, 192f, 196f, 199f, 202f,
            205f, 209f, 212f, 215f, 218f, 222f, 225f, 228f, 231f, 235f, 238f,
            241f, 244f, 247f, 251f, 254f, 257f, 260f, 264f, 267f
    };

    //ENCAPSULATION:

    //CONSTRUCTOR
    Radius (){
//       super.upBound = 9;
        super.NAME1="лучевая кость";
        super.NAME2="лучевой кости";
        super.df = new int[upBound];
        super.method = new String[upBound][2];
    }



    //ОПРЕДЕЛЕНИЕ ПОЛА:
    @Override
    /* определение диагностических коэффициентов */
    public String getDF(float  measurement, int step){


        //else -> сообщение об ошибке!
        return "ДК = "+method[step][1];
    };



}
