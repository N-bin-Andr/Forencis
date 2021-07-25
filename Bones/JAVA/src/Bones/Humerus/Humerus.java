package Bones.Humerus;

import Bones.Bones;

public class Humerus extends Bones {

//FIELDS:

    //ARRAYS:
    //MANOUVRIER
    private static final float[]mLenght_Manouvrier = {
            //длина плечевой кости мужчин  (рост 153-183см).
            295f, 298f, 302f, 306f, 309f, 313f, 316f, 320f, 324f, 328f, 332f,
            336f, 340f, 344f, 348f, 352f, 356f, 360f, 364f, 368f
    };
    private static final float[]wLenght_Manouvrier = {
            //длина плечевой кости женщин (рост 140-171,5см).
            263f, 266f, 270f, 273f, 276f, 279f, 282f, 285f, 289f, 292f, 297f,
            302f, 307f, 313f, 318f, 324f, 329f, 334f, 339f, 344f
    };

    //TROTTER_&_GLESER
    private float[]mGrPr_TrotterGleser;//рост мужчины.
    private float[]wGrPr_TrotterGleser;//рост женщины.
    private static final float [] mL_TROTTER_GLESER={
            //длина мужской плечевой кости (рост 152-198см).
            265f, 268f, 271f, 275f, 278f, 281f, 284f, 288f, 291f, 294f, 297f,
            301f, 304f, 307f, 310f, 314f, 317f, 320f, 323f, 327f, 330f, 333f,
            336f, 339f, 343f, 346f, 349f, 352f, 356f, 359f, 362f, 365f, 369f,
            372f, 375f, 378f, 382f, 385f, 388f, 391f, 395f, 398f, 401f, 404f,
            408f, 411f, 414f

/* по Алексееву (1966) на основании TrotterGleser
    длина плечевой кости мужчин (негроидов) (рост 152-198см.):
            266f, 269f, 273f, 276f, 280f, 283f, 287f, 290f, 294f, 297f, 300f, 304f,
            307f, 311f, 314f, 318f, 321f, 325f, 328f, 332f, 335f, 339f, 342f, 346f,
            349f, 353f, 356f, 359f, 363f, 366f, 370f, 373f, 377f, 380f, 384f, 387f,
            391f, 394f, 398f, 401f, 405f, 408f, 412f, 415f, 418f, 422f, 425f

   длина плечевой кости мужчин (европеоидов) (рост 152-198см.):
            256f, 259f, 263f, 266f, 270f, 273f, 276f, 280f, 283f, 287f, 290f, 294f,
            297f, 301f, 304f, 308f, 311f, 315f, 318f, 321f, 325f, 328f, 332f, 335f,
            339f, 342f, 346f, 349f, 353f, 356f, 360f, 363f, 366f, 370f, 373f, 377f,
            380f, 384f, 387f, 390f, 394f, 398f, 401f, 404f, 411f, 415f

    длина плечевой кости мужчин (монголоидов) (рост 152-198см.):
            257f, 260f, 264f, 268f, 272f, 276f, 280f, 283f, 287f, 290f, 294f, 298f,
            302f, 305f, 309f, 313f, 316f, 320f, 324f, 328f, 331f, 335f, 339f, 343f,
            346f, 350f, 354f, 358f, 361f, 365f, 369f, 372f, 376f, 380f, 384f, 387f,
            391f, 395f, 399f, 402f, 406f, 410f, 413f, 417f, 421f, 425f, 428f
 */

    } ;
    private static final float[]wL_TROTTER_GLESER={
            //длина плечевой кости женщин (европеоиды) (рост 140-184см):
            244f, 247f, 250f, 253f, 256f, 259f, 262f, 265f, 268f, 271f, 274f,
            277f, 280f, 283f, 286f, 289f, 292f, 295f, 298f, 301f, 304f, 307f,
            310f, 313f, 316f, 319f, 322f, 324f, 327f, 330f, 333f, 336f, 339f,
            342f, 345f, 348f, 351f, 354f, 357f, 360f, 363f, 366f, 369f, 372f,
            375f
    };
    private static final float[]wnL_TROTTER_GLESER={
            //длина плечевой кости женщин (негроиды) (рост 140-184см):
            245f, 248f, 251f, 254f, 258f, 261f, 264f, 267f, 271f, 274f, 277f,
            280f, 284f, 287f, 290f, 293f, 297f, 300f, 303f, 306f, 310f, 313f,
            316f, 319f, 322f, 326f, 329f, 332f, 335f, 339f, 342f, 345f, 348f,
            352f, 355f, 358f, 361f, 365f, 368f, 371f, 374f, 378f, 381f, 384f,
            384f
    };

    //TELCCA
    private float[]mGrPr_Telcca;//рост мужчины.
    private float[]wGrPr_Telcca;//рост женщины.
    private static final  float [] mL_TELCCA={
    //длина мужской кости.
            278f, 281f, 285f, 288f, 292f, 296f, 299f, 303f, 306f, 310f, 313f,
            317f, 320f, 324f, 328f, 331f, 335f, 338f, 342f, 346f, 349f, 353f,
            356f, 360f, 363f, 367f, 371f, 374f, 378f, 381f, 385f
    };
    private static final float [] wL_TELCCA = {
    //длина женской кости.
            263f, 267f, 271f, 274f, 278f, 282f, 285f, 289f, 293f, 297f, 300f,
            304f, 308f, 312f, 315f, 319f, 323f, 326f, 330f, 334f, 337f, 341f,
            345f, 348f, 352f, 356f, 360f, 363f, 367f, 371f, 374f
    };

    //ENCAPSULATION:

    //CONSTRUCTOR
    Humerus (){
        super.upBound = 9;
        super.NAME1="плечевая кость";
        super.NAME2="плечевой кости";
        super.df = new int[upBound];
        super.method = new String[upBound][2];
    }
//measurement

//ОПРЕДЕЛЕНИЕ ПОЛА:
 @Override
 /* определение диагностических коэффициентов */
    public String getDF(float  measurement, int step){
        method[step][1]="НПВ";
        if (measurement>0f) {
            switch (step) {
                case 0: {
                    method[step][0] = "Наибольшая длина кости в естественном положении (мм.)";
                    if (measurement <= 283f) method[step][1] = "+беск.";
                    else if (measurement > 353f) method[step][1] = "-беск";
                    else {
                        if (measurement > 283f & measurement <= 303f) method[step][1] = "+159";
                        else if (measurement > 303f & measurement <= 313f) method[step][1] = "+61";
                        else if (measurement > 313f & measurement <= 323f) method[step][1] = "+23";
                        else if (measurement > 323f & measurement <= 343f) method[step][1] = "-59";
                        else if (measurement > 343f & measurement <= 353f) method[step][1] = "-128";
                        df[step] = Integer.parseInt(method[step][1]);
                    }
                    break;}

                case 1: {
                    method[step][0] = "Окружность в середине диафиза (мм.)";
                    if (measurement <= 57f) method[1][1] = "+беск.";
                    else if (measurement > 78f) method[1][1] = "-беск.";
                    else {
                        if (measurement > 57f & measurement <= 60f) method[1][1] = "+151";
                        else if (measurement > 60f & measurement <= 63f) method[1][1] = "+86";
                        else if (measurement > 63f & measurement <= 66f) method[1][1] = "-+12";
                        else if (measurement > 66f & measurement <= 69f) method[1][1] = "-47";
                        else if (measurement > 69f & measurement <= 72f) method[1][1] = "-97";
                        else if (measurement > 72f & measurement <= 78f) method[1][1] = "-116";
                        df[step] = Integer.parseInt(method[1][1]);
                    }
                    break;}

                case 2: {
                    method[step][0] = "Минимальная окружность диафиза (мм.)";
                    if (measurement <= 52f) method[2][1] = "+беск.";
                    else if (measurement > 70f) method[2][1] = "-беск.";
                    else {
                        if (measurement > 52f & measurement <= 58f) method[2][1] = "+175";
                        else if (measurement > 58f & measurement <= 61f) method[2][1] = "+78";
                        else if (measurement > 61f & measurement <= 64f) method[2][1] = "-21";
                        else if (measurement > 64f & measurement <= 67f) method[2][1] = "-76";
                        else if (measurement > 67f & measurement <= 70f) method[2][1] = "-148";
                        df[step] = Integer.parseInt(method[2][1]);
                    }
                    break;}

                case 3: {
                    method[step][0] = "Окружность головки (мм.)";
                    if (measurement > 0f) {
                        if (measurement <= 124f) method[3][1] = "+беск.";
                        else if (measurement > 154f) method[3][1] = "-беск.";
                        else {
                            if (measurement > 124f & measurement <= 129f) method[3][1] = "+159";
                            else if (measurement > 129f & measurement <= 134f) method[3][1] = "+120";
                            else if (measurement > 134f & measurement <= 139f) method[3][1] = "-+10";
                            else if (measurement > 139f & measurement <= 144f) method[3][1] = "-35";
                            else if (measurement > 144f & measurement <= 154f) method[3][1] = "-176";
                            df[step] = Integer.parseInt(method[3][1]);
                        }

                    }
                    break;}

                case 4: {
                    method[step][0] = "Ширина дистального эпифиза (мм.)";
                    if (measurement <= 52f) method[4][1] = "+беск.";
                    else if (measurement > 70f) method[4][1] = "-беск.";
                    else {
                        if (measurement > 52f & measurement <= 55f) method[4][1] = "+125";
                        else if (measurement > 55f & measurement <= 58f) method[4][1] = "+84";
                        else if (measurement > 58f & measurement <= 61f) method[4][1] = "+7";
                        else if (measurement > 61f & measurement <= 64f) method[4][1] = "-55";
                        else if (measurement > 64f & measurement <= 70f) method[4][1] = "-137";
                        df[step] = Integer.parseInt(method[4][1]);
                    }
                    break;}

                case 5: {
                    method[step][0] = "Площадь компактного вещества на поперечном распиле середины диафиза (мм.)";
                    if (measurement <= 159f) method[5][1] = "+99";
                    else if (measurement > 259f) method[5][1] = "-беск";
                    else {
                        if (measurement > 159f & measurement <= 179f) method[5][1] = "+87";
                        else if (measurement > 179f & measurement <= 199f) method[5][1] = "+56";
                        else if (measurement > 199f & measurement <= 219f) method[5][1] = "-14";
                        else if (measurement > 219f & measurement <= 259f) method[5][1] = "-163";
                        df[step] = Integer.parseInt(method[5][1]);
                    }
                    break;}

                case 6: {
                    method[step][0] = "Площадь поперечного распила середины диафиза (мм.)";
                    if (measurement <= 200f) method[6][1] = "+беск.";
                    else if (measurement > 380f) method[6][1] = "-беск";
                    else {
                        if (measurement > 200f & measurement <= 240f) method[6][1] = "+110";
                        else if (measurement > 240f & measurement <= 280f) method[6][1] = "+70";
                        else if (measurement > 280f & measurement <= 300f) method[6][1] = "+13";
                        else if (measurement > 300f & measurement <= 320f) method[6][1] = "-22";
                        else if (measurement > 320f & measurement <= 340f) method[6][1] = "-68";
                        else if (measurement > 340f & measurement <= 360f) method[6][1] = "-83";
                        else if (measurement > 360f & measurement <= 380f) method[6][1] = "-123";
                        df[step] = Integer.parseInt(method[6][1]);
                    }
                    break;}

                case 7: {
                    method[step][0] = "Минимальный диаметр диафиза (мм.)";
                    if (measurement <= 15f) method[7][1] = "+беск.";
                    else if (measurement > 21f) method[7][1] = "-беск";
                    else {
                        if (measurement > 15f & measurement <= 17f) method[7][1] = "+172";
                        else if (measurement > 17f & measurement <= 18f) method[7][1] = "+75";
                        else if (measurement > 18f & measurement <= 19f) method[7][1] = "-38";
                        else if (measurement > 19f & measurement <= 21f) method[7][1] = "-124";
                        df[step] = Integer.parseInt(method[7][1]);
                    }
                    break;}

                case 8: {
                    method[step][0] = "Толщина компактного вещества в области минимального диаметра диафиза (мм.)";
                    if (measurement <= 4f) method[8][1] = "+беск.";
                    else if (measurement > 12f) method[8][1] = "-беск";
                    else {
                        if (measurement > 4f & measurement <= 5f) method[8][1] = "+112";
                        else if (measurement > 5f & measurement <= 7f) method[8][1] = "+61";
                        else if (measurement > 7f & measurement <= 9f) method[8][1] = "0";
                        else if (measurement > 9f & measurement <= 11f) method[8][1] = "-60";
                        else if (measurement > 11f & measurement <= 12f) method[8][1] = "-118";
                        df[step] = Integer.parseInt(method[8][1]);
                    }
                    break;}
            }
        }
        //else -> сообщение об ошибке!
        return "ДК = "+method[step][1];
    }

}
