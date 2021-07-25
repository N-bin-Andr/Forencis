package Bones.Fibula;

import Bones.Bones;

public class Fibula extends Bones {


    //FIELDS:

    //ARRAYS:
    //MANOUVRIER
    private float[]mLenght_Manouvrier = {
    //длина малоберцовой кости мужчин (рост 153-183см):
            318f, 323f, 328f, 333f, 338f, 344f, 349f, 353f, 358f, 363f, 368f,
            373f, 378f, 383f, 388f, 393f, 398f, 403f, 408f, 413f
    };

    private float[]wLenght_Manouvrier = {
            //длина малобероцовой кости женщин (рост 140-171,5см).
            283f, 288f, 293f, 298f, 303f, 307f, 311f, 316f, 320f, 325f, 330f,
            336f, 341f, 346f, 351f, 356f, 361f, 366f, 371f, 376f
    };

    //TROTTER_&_GLESER
    private float[]mGrPr_TrotterGleser;//рост мужчины.
    private float[]wGrPr_TrotterGleser;//рост женщины.
    private static final float [] mL_TROTTER_GLESER={
    //длина малоберцовой кости мужчин (рост 152-198см.):
            299f, 303f, 307f, 311f, 314f, 318f, 322f, 326f, 329f, 333f, 337f,
            340f, 344f, 348f, 352f, 355f, 359f, 363f, 367f, 370f, 374f, 378f,
            381f, 385f, 389f, 393f, 396f, 400f, 404f, 408f, 411f, 415f, 419f,
            422f, 426f, 430f, 434f, 437f, 441f, 445f, 449f, 452f, 456f, 460f,
            463f, 467f, 471f
    } ;

/*по Алексееву на основании TrotterGleser
    длина малоберцовой кости мужчин (негроиды) (рост 152-198см.):
            307f, 312f, 316f, 320f, 324f, 329f, 333f, 337f, 342f, 346f, 350f, 354f,
            359f, 363f, 367f, 371f, 376f, 380f, 384f, 389f, 393f, 397f, 401f, 406f,
            410f, 414f, 419f, 423f, 427f, 431f, 436f, 440f, 444f, 448f, 453f, 457f,
            461f, 466f, 470f, 474f, 478f, 483f, 487f, 491f, 495f, 500f, 504f

    длина малоберцовой кости мужчин (европеоиды) (рост 152-198см.):
            294f, 298f, 302f, 306f, 310f, 313f, 317f, 321f, 325f, 329f, 333f, 337f,
            340f, 344f, 348f, 352f, 356f, 360f, 363f, 367f, 371f, 375f, 379f, 383f,
            387f, 390f, 394f, 398f, 402f, 406f, 410f, 413f, 417f, 421f, 425f, 429f,
            433f, 436f, 440f, 444f, 448f, 452f, 456f, 460f, 463f, 467f, 471

    длина малоберцовой кости мужчин (монголоиды) (рост 152-198см.):
            298f, 302f, 306f, 310f, 314f, 318f, 323f, 327f, 331f, 335f, 339f, 343f,
            348f, 352f, 356f, 360f, 364f, 368f, 373f, 377f, 381f, 385f, 389f, 393f,
            398f, 402f, 406f, 410f, 414f, 418f, 423f, 427f, 431f, 435f, 439f, 443f,
            448f, 452f, 456f, 460f, 464f, 468f, 473f, 477f, 481f, 485f, 489f
 */

    private static final float[]wL_TROTTER_GLESER={
            //длина малоберцовой кости женщин (европеоиды) (рост 140-184см.):
            274f, 278f, 281f, 285f, 288f, 291f, 295f, 298f, 302f, 305f, 309f,
            312f, 315f, 319f, 322f, 326f, 329f, 332f, 336f, 340f, 343f, 346f,
            349f, 353f, 356f, 360f, 363f, 366f, 370f, 373f, 377f, 380f, 384f,
            387f, 390f, 394f, 397f, 401f, 404f, 407f, 411f, 414f, 418f, 421f,
            425f
    };

    private static final float[]wnL_TROTTER_GLESER={
            //длина малоберцовой кости женщин  (негроиды) (рост 140-184см.):
            278f, 282f, 286f, 290f, 294f, 298f, 302f, 306f, 310f, 314f, 318f,
            322f, 326f, 330f, 334f, 338f, 342f, 346f, 350f, 354f, 358f, 362f,
            366f, 370f, 374f, 378f, 382f, 386f, 390f, 394f, 398f, 402f, 406f,
            410f, 414f, 418f, 422f, 426f, 430f, 434f, 438f, 442f, 446f, 450f,
            454f
    };

    //TELCCA
    private float[]mGrPr_Telcca;//рост мужчины.
    private float[]wGrPr_Telcca;//рост женщины.
    private static final  float [] mL_TELCCA={
    //длина малоберцовой кости мужчин (рост 155-185см).
            303f, 307f, 311f, 315f, 319f, 323f, 327f, 331f, 335f, 339f, 343f,
            348f, 352f, 356f, 360f, 364f, 368f, 372f, 376f, 380f, 384f, 388f,
            392f, 396f, 400f, 404f, 408f, 412f, 416f, 420f, 424f
    };
    private  static final float [] wL_TELCCA = {
    //длина малоберцовой кости женщин (рост 145-175см).
            276f, 280f, 284f, 289f, 293f, 298f, 302f, 306f, 311f, 315f, 320f,
            324f, 328f, 332f, 337f, 341f, 345f, 350f, 354f, 358f, 363f, 367f,
            372f, 376f, 381f, 385f, 389f, 394f, 398f, 403f, 407f

    };

    //ENCAPSULATION:

    //CONSTRUCTOR
    Fibula (){
        super.upBound = 2;
        super.NAME1="малоберцовая кость";
        super.NAME2="малоберцовой кости";
        super.df = new int[upBound];
        super.method = new String[upBound][2];

    }

//ОПРЕДЕЛЕНИЕ ПОЛА:

    @Override
    public String getDF(float  measurement, int step){
    //определение диагностических коэффициентов.
        if (measurement>0){
            method[step][1]="НПВ";
            switch (step){
                case 0:{
                    method[step][0]="Наибольшая длина кости (мм.)";
                    if (measurement<=310f)method[step][1]="+беск.";
                    else if (measurement>400f)method[step][1]="-беск";
                    else {
                        if (measurement > 310f & measurement <= 330f) method[step][1] = "+132";
                        else if (measurement > 330f & measurement <= 340f) method[step][1] = "+92";
                        else if (measurement > 340f & measurement <= 350f) method[step][1] = "+16";
                        else if (measurement > 350f & measurement <= 360f) method[step][1] = "-13";
                        else if (measurement > 360f & measurement <= 380f) method[step][1] = "-81";
                        else if (measurement > 380f & measurement <= 400f) method[step][1] = "-129";
                        df[step] = Integer.parseInt(method[step][1]);
                    }
                break;
                }
                case 1:{
                    method[step][0]="Ширина проксимального эпифиза";
                    if (measurement<=22f)method[step][1]="-беск.";
                    else if (measurement>33f)method[step][1]="+беск.";
                    else {
                        if (measurement>22f & measurement<=25f)method[step][1]="+60";
                        else if (measurement>25f & measurement<=27f)method[step][1]="+48";
                        else if (measurement>27f & measurement<=28f)method[step][1]="+0";
                        else if (measurement>28f & measurement<=30f)method[step][1]="-38";
                        else if (measurement>30f & measurement<=33f)method[step][1]="-132";
                        df[step] = Integer.parseInt(method[step][1]);
                    }
                break;
                }
            }
        }
        //else -> сообщение об ошибке!
        return "ДК = "+method[step][1];
    }


    private static void main (String[]args){
    //КЛИЕНТ
        Fibula fib = new Fibula();
        int i =0;
            for (i=0; i>2; i++ );{

                System.out.println(fib.getDF(330,i));
        }





    }


//    //определение диагностических коэффициентов
//    public String getDF(float  measurement, int i){
//        //1). Max. Length
//        method[i][0]="Наибольшая длина кости (мм.)";
//        method[i][1]="НПВ";
//        if (measurement>0f)
//        {
//            if (measurement<=310f)method[i][1]="+беск.";
//            else if (measurement>400f)method[i][1]="-беск";
//            else {
//                if (measurement > 310f & measurement <= 330f) method[i][1] = "+132";
//                else if (measurement > 330f & measurement <= 340f) method[i][1] = "+92";
//                else if (measurement > 340f & measurement <= 350f) method[i][1] = "+16";
//                else if (measurement > 350f & measurement <= 360f) method[i][1] = "-13";
//                else if (measurement > 360f & measurement <= 380f) method[i][1] = "-81";
//                else if (measurement > 380f & measurement <= 400f) method[i][1] = "-129";
//                df[0] = Integer.parseInt(method[i][1]);
//            }
//        }
//        //else -> сообщение об ошибке!
//        return "ДК = "+method[i][1];
//    }
//
//    public String getDF(float  measurement, int i){
//        //2).Widgth Pr. Epifise
//        method[i][i]="Ширина проксимального эпифиза";
//        method[i][1]="НПВ";
//        if (measurement>0f){
//            if (measurement<=22f)method[i][1]="-беск.";
//            else if (measurement>33f)method[i][1]="+беск.";
//            else {
//                if (measurement>22f & measurement<=25f)method[i][1]="+60";
//                else if (measurement>25f & measurement<=27f)method[i][1]="+48";
//                else if (measurement>27f & measurement<=28f)method[i][1]="+0";
//                else if (measurement>28f & measurement<=30f)method[i][1]="-38";
//                else if (measurement>30f & measurement<=33f)method[i][1]="-132";
//                df[1] = Integer.parseInt(method[i][1]);
//            }
//        }
//        //else -> сообщение об ошибке!
//        return "ДК = "+method[i][1];
//    }

}
