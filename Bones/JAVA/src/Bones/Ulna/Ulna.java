 package Bones.Ulna;

 import Bones.Bones;

 public class Ulna extends Bones {



    //ARRAYS:
    //MANOUVRIER
    private float[]MANOUVRIER_mLenght = {
            //длина локтевой кости мужчин (рост 153-183см)
            227f, 231f, 235f, 239f, 243f, 246f, 249f, 253f, 257f, 260f, 263f,
            266f, 270f, 273f, 276f, 280f, 283f, 287f, 290f, 293f
    };

    private float[]MANOUVRIER_wLengh = {
            //длина локтевой кости женщин (рост 140-171,5см).
            203f, 206f, 209f, 212f, 215f, 217f, 219f, 222f, 225f, 228f, 231f,
            235f, 239f, 243f, 247f, 251f, 254f, 258f, 261f, 264f
    };

    //TROTTER_&_GLESER
    private float[]mGrPr_TrotterGleser;//рост мужчины.
    private float[]wGrPr_TrotterGleser;//рост женщины.

    private static final float [] mL_TROTTER_GLESER={
            //длина локтевой кости мужчин (рост 152-198см.):
            211f, 213f, 216f, 219f, 222f, 224f, 227f, 230f, 232f, 235f, 238f,
            240f, 243f, 246f, 249f, 251f, 254f, 257f, 259f, 262f, 265f, 267f,
            270f, 273f, 276f, 278f, 281f, 284f, 286f, 289f, 292f, 294f, 297f,
            300f, 303f, 305f, 308f, 311f, 313f, 316f, 319f, 321f, 324f, 327f,
            330f, 332f, 335f
/*
    по Алексееву (1966) на основании TrotterGleser
    длина локтевой кости мужчин (негроидов) (рост 152-198см.):
            216f, 219f, 223f, 226f, 229f, 232f, 235f, 238f, 241f, 244f, 248f, 251f,
            254f, 257f, 260f, 263f, 266f, 269f, 273f, 276f, 279f, 282f, 285f, 288f,
            291f, 294f, 298f, 301f, 304f, 307f, 310f, 313f, 316f, 319f, 323f, 327f,
            329f, 332f, 335f, 338f, 341f, 344f, 348f, 351f, 354f, 357f, 360f

     длина локтевой кости мужчин (европеоиды) (рост 152-198см.):
            203f, 206f, 209f, 211f, 214f, 217f, 219f, 222f, 225f, 227f, 230f, 233f,
            235f, 238f, 241f, 243f, 246f, 249f, 251f, 254f, 257f, 259f, 262f, 264f,
            267f, 270f, 272f, 275f, 278f, 280f, 283f, 286f, 288f, 291f, 294f, 296f,
            299f, 302f, 304f, 307f, 310f, 312f, 315f, 318f, 320f, 323f, 326f

     длина локтевой кости мужчин (монголоиды) (рост 152-198см.):
            214f, 217f, 220f, 223f, 226f, 229f, 231f, 234f, 237f, 240f, 243f, 246f,
            249f, 252f, 254f, 257f, 260f, 263f, 266f, 269f, 272f, 275f, 277f, 280f,
            283f, 286f, 289f, 292f, 295f, 298f, 300f, 303f, 306f, 309f, 312f, 315f,
            318f, 321f, 323f, 326f, 329f, 332f, 335f, 338f, 341f, 344f, 346f
 */
    } ;
    private static final float[]wL_TROTTER_GLESER={
            //длина женской кости.
            193f, 195f, 197f, 200f, 202f, 204f, 207f, 209f, 211f, 214f, 216f,
            218f, 221f, 223f, 225f, 228f, 230f, 232f, 235f, 237f, 239f, 242f,
            244f, 246f, 249f, 251f, 253f, 256f, 258f, 261f, 263f, 265f, 268f,
            270f, 272f, 275f, 277f, 279f, 282f, 284f, 286f, 289f, 291f, 293f,
            296f
    };
    private static final float[]wnL_TROTTER_GLESER={
            //длина женской кости (негроиды).
            195f, 198f, 201f, 204f, 207f, 210f, 213f, 216f, 219f, 222f, 225f,
            228f, 231f, 235f, 238f, 241f, 244f, 247f, 250f, 253f, 256f, 259f,
            262f, 265f, 268f, 271f, 274f, 277f, 280f, 283f, 286f, 289f, 292f,
            295f, 298f, 301f, 304f, 307f, 310f, 313f, 316f, 319f, 322f, 325f,
            328f
    };

    //TELCCA
    private float[]mGrPr_Telcca;//рост мужчины.
    private float[]wGrPr_Telcca;//рост женщины.
    private static final  float [] mL_TELCCA={
    //длина мужской кости.
            186f, 189f, 192f, 195f, 198f, 202f, 205f, 208f, 211f, 214f, 217f,
            220f, 224f, 227f, 230f, 233f, 236f, 239f, 242f, 245f, 249f, 252f,
            255f, 258f, 261f, 264f, 267f, 270f, 274f, 277f, 280f
    };
    private static final float [] wL_TELCCA = {
    //длина женской кости.
            177f, 180f, 183f, 186f, 189f, 192f, 195f, 198f, 202f, 205f, 208f,
            211f, 214f, 217f, 220f, 223f, 226f, 229f, 232f, 235f, 238f, 241f,
            244f, 247f, 250f, 253f, 256f, 259f, 262f, 265f, 268f
    };

    //ENCAPSULATION:

    //CONSTRUCTOR
    Ulna (){
//        super.upBound = 9;
        super.NAME1="локтевая кость";
        super.NAME2="локтевой кости";
        super.df = new int[upBound];
        super.method = new String[upBound][];
    }



    //ОПРЕДЕЛЕНИЕ ПОЛА:
    @Override
    /* определение диагностических коэффициентов */
    public String getDF(float  measurement, int step){


        //else -> сообщение об ошибке!
        return "ДК = "+method[step][1];
    };
}
