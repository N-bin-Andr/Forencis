package Bones.Femoris;

import Bones.Bones;

public class Femoris extends Bones {

    //FIELDS:

    //ARRAYS:
    //MANOUVRIER
    private float[]mLenght_Manouvrier = {
        //длина бедренной кости мужчин (рост 153-183см):
       392f, 398f, 404f, 410f, 416f, 422f, 428f, 434f, 440f, 446f, 453f,
       460f, 467f, 475f, 482f, 490f, 497f, 504f, 512f, 519f
    };

    private float[]wLenght_Manouvrier = {
        //длина бедренной кости женщин (рост 140-171,5см):
        363f, 368f, 373f, 378f, 383f, 388f, 393f, 398f, 403f, 408f, 415f,
        422f, 429f, 436f, 443f, 450f, 457f, 464f, 471f, 478f
    };


    //TROTTER_&_GLESER
    private float[]mGrPr_TrotterGleser;//рост мужчины.
    private float[]wGrPr_TrotterGleser;//рост женщины.
    private static final float [] mL_TROTTER_GLESER={
    //длина бедренной кости мужщин (рост 152-198см):
            381f, 385f, 389f, 393f, 398f, 402f, 406f, 410f, 414f, 419f, 423f,
            427f, 431f, 435f, 440f, 444f, 448f, 452f, 456f, 461f, 465f, 469f,
            473f, 477f, 482f, 486f, 490f, 494f, 498f, 503f, 507f, 511f, 515f,
            519f, 524f, 528f, 532f, 536f, 540f, 545f, 549f, 553f, 557f, 561f,
            566f, 570f, 574f

/*            по Алексееву (TrotterGleser)
     длина бедренной кости  мужчин (европеоидов) (рост 152-198см):
            373f, 377f, 381f, 386f, 390f, 394f, 399f, 403f, 407f, 412f, 416f,
            420f, 424f, 429f, 433f, 437f, 442f, 446f, 450f, 455f, 459f, 463f,
            468f, 472f, 476f, 780f, 485f, 489f, 493f, 498f, 502f, 506f, 511f,
            515f, 519f, 524f, 528f, 532f, 537f, 541f, 545f, 549f, 554f, 558f,
            563f, 567f, 571f

    длина бедренной кости  мужчин (монголоидов) (рост 152-198см):
            369f, 374f, 379f, 383f, 388f, 393f, 397f, 402f, 407f, 411f, 416f,
            421f, 425f, 430f, 434f, 439f, 444f, 449f, 453f, 458f, 462f, 467f,
            472f, 476f, 481f, 486f, 490f, 495f, 500f, 504f, 509f, 514f, 518f,
            523f, 528f, 532f, 537f, 542f, 546f, 551f, 555f, 560f, 565f, 569f,
            574f, 579f, 583f

     длина бедренной кости мужчин (негроидов) (рост 152-198см):
            380f, 385f, 389f, 394f, 399f, 404f, 408f, 413f, 418f, 423f, 428f, 432f,
            437f, 442f, 447f, 451f, 456f, 461f, 466f, 470f, 475f, 480f, 485f, 489f,
            494f, 499f, 504f, 508f, 513f, 518f, 523f, 528f, 532f, 537f, 542f, 547f,
            551f, 556f, 561f, 566f, 570f, 575f, 580f, 585f, 589f, 594f, 599f
 */
} ;
    private static final float[]wL_TROTTER_GLESER={
    //длина бедренной кости женщин (европеоиды) (рост 140-184см):
            348f, 352f, 356f, 360f, 364f, 368f, 372f, 376f, 380f, 384f, 388f,
            392f, 396f, 400f, 404f, 409f, 413f, 417f, 421f, 425f, 429f, 433f,
            437f, 441f, 445f, 449f, 453f, 457f, 461f, 465f, 469f, 473f, 477f,
            481f, 485f, 489f, 494f, 498f, 502f, 506f, 510f, 514f, 518f, 522f,
            526f
    };
    private static final float[]wnL_TROTTER_GLESER={
            //длина бедренной кости женщин  (негроиды) (рост 140-184см):
            352f, 356f, 361f, 365f, 369f, 374f, 378f, 383f, 387f, 391f, 396f,
            400f, 405f, 409f, 413f, 418f, 422f, 426f, 431f, 435f, 440f, 444f,
            448f, 453f, 457f, 462f, 466f, 470f, 475f, 479f, 484f, 488f, 492f,
            497f, 501f, 505f, 510f,  514f, 519f, 523f, 527f, 532f, 536f, 541f,
            545f
    };

    //TELCCA
    private static final  float [] mL_TELCCA={
    //
            387f, 391f, 396f, 401f, 406f, 410f, 415f, 420f, 425f, 430f, 434f, 439f,
            444f, 448f, 453f, 458f, 463f, 468f, 472f, 477f, 482f, 487f, 492f, 496f,
            501f, 506f, 511f, 515f, 520f, 525f, 529f
    };
    private static final float [] wL_TELCCA = {
    //длина женской кости.
            352f, 357f, 363f, 369f, 375f, 380f, 386f, 392f, 397f, 403f, 408f, 414f,
            419f, 425f, 430f, 436f, 441f, 447f, 453f, 458f, 463f, 469f, 474f, 480f,
            485f, 491f, 496f, 502f, 508f, 513f, 518f
    };
    private float[]mGrPr_Telcca;//рост мужчины.
    private float[]wGrPr_Telcca;//рост женщины.

//  CONSTRUCTOR
   private Femoris(){
       super.upBound=9;
       super.NAME1="бедренная кость";
       super.NAME2="бедренной кости";
       super.df = new int[upBound];
       super.method = new String[upBound][2];
    }

    //ОПРЕДЕЛЕНИЕ ПОЛА:
    @Override
    public String getDF(float measurement, int step) {
    //определение диагностических коэффициентов
        method[step][1] = "НПВ";
        if (measurement>0f) {
            switch (step) {
                case 0: {
                    method[step][0] = "Длина тела бедренной кости в естественном положении (мм.)";
                    if (measurement > 43f) {//max размер у мальчиков 14-15 лет{
                        if (measurement <= 392f) method[step][1] = "+беск.";
                        else if (measurement > 482f) method[step][1] = "-беск";
                        else {
                            if (measurement > 392f & measurement <= 412f) method[step][1] = "+151";
                            else if (measurement > 412f & measurement <= 422f) method[step][1] = "+56";
                            else if (measurement > 422f & measurement <= 442f) method[step][1] = "+2";
                            else if (measurement > 442f & measurement <= 452f) method[step][1] = "-31";
                            else if (measurement > 452f & measurement <= 462f) method[step][1] = "-79";
                            else if (measurement > 462f & measurement <= 482f) method[step][1] = "-111";
                            df[step] = Integer.parseInt(method[step][1]);
                        }
                    } else
                        method[step][1] = "Длина кости менее 44см! Исследуемая бедренная кость может принадлежать ребенку!";
                    break;
                }

                case 1: {
                    method[step][0] = "Окружность в середине диафиза (мм.)";
                    if (measurement <= 73f) method[step][1] = "+беск.";
                    else if (measurement > 97f) method[step][1] = "-беск";
                    else {
                        if (measurement > 73f & measurement <= 79f) method[step][1] = "+103";
                        else if (measurement > 79f & measurement <= 82f) method[step][1] = "+78";
                        else if (measurement > 82f & measurement <= 85f) method[step][1] = "+35";
                        else if (measurement > 85f & measurement <= 88f) method[step][1] = "-12";
                        else if (measurement > 88f & measurement <= 91f) method[step][1] = "-34";
                        else if (measurement > 91f & measurement <= 97f) method[step][1] = "-159";
                        df[step] = Integer.parseInt(method[step][1]);
                    }
                    break;
                }

                case 2: {
                    method[step][0] = "Окружность головки (мм.)";
                    if (measurement <= 130f) method[step][1] = "+беск.";
                    else if (measurement > 160f) method[step][1] = "-беск";
                    else {
                        if (measurement > 130f & measurement <= 135f) method[step][1] = "+115";
                        else if (measurement > 135f & measurement <= 140f) method[step][1] = "+73";
                        else if (measurement > 140f & measurement <= 145f) method[step][1] = "+35";
                        else if (measurement > 145f & measurement <= 150f) method[step][1] = "-40";
                        df[step] = Integer.parseInt(method[step][1]);
                    }
                    break;
                }

                case 3: {
                    method[step][0] = "Ширина дистального эпифиза (мм.)";
                    if (measurement <= 72f) method[step][1] = "+беск.";
                    else if (measurement > 90f) method[step][1] = "-беск";
                    else {
                        if (measurement > 72f & measurement <= 75f) method[step][1] = "+128";
                        else if (measurement > 75f & measurement <= 78f) method[step][1] = "+68";
                        else if (measurement > 78f & measurement <= 81f) method[step][1] = "-12";
                        else if (measurement > 81f & measurement <= 84f) method[step][1] = "-97";
                        else if (measurement > 84f & measurement <= 90f) method[step][1] = "-176";
                        df[step] = Integer.parseInt(method[step][1]);
                    }
                    break;
                }

                case 4: {
                    method[step][0] = "Степень изгиба (мм.)";
                    if (measurement <= 48f) method[step][1] = "+беск.";
                    else if (measurement > 69f) method[step][1] = "-беск";
                    else {
                        if (measurement > 48f & measurement <= 54f) method[step][1] = "+110";
                        else if (measurement > 54f & measurement <= 57f) method[step][1] = "+4";
                        else if (measurement > 57f & measurement <= 60f) method[step][1] = "-9";
                        else if (measurement > 60f & measurement <= 63f) method[step][1] = "-33";
                        else if (measurement > 63f & measurement <= 66f) method[step][1] = "-71";
                        else if (measurement > 66f & measurement <= 69f) method[step][1] = "-101";
                        df[step] = Integer.parseInt(method[step][1]);
                    }

                    break;
                }

                case 5: {
                    method[step][0] = "Площадь компактного вещества на поперечном распиле середины диафиза (мм.)";
                    if (measurement <= 274f) method[step][1] = "+беск.";
                    else if (measurement > 514f) method[step][1] = "-беск";
                    else {
                        if (measurement > 274 & measurement <= 334) method[step][1] = "+120";
                        else if (measurement > 334f & measurement <= 364f) method[step][1] = "+44";
                        else if (measurement > 364f & measurement <= 394f) method[step][1] = "+3";
                        else if (measurement > 394f & measurement <= 424f) method[step][1] = "-29";
                        else if (measurement > 424f & measurement <= 454f) method[step][1] = "-61";
                        else if (measurement > 454f & measurement <= 514f) method[step][1] = "-117";
                        df[step] = Integer.parseInt(method[step][1]);
                    }
                    break;
                }

                case 6: {
                    method[step][0] = "Площадь поперечного распила середины диафиза (мм.)";
                    if (measurement <= 414f) method[step][1] = "+беск.";
                    else if (measurement > 654f) method[step][1] = "-беск";
                    else {
                        if (measurement > 414 & measurement <= 474) method[step][1] = "+112";
                        else if (measurement > 474f & measurement <= 504f) method[step][1] = "+44";
                        else if (measurement > 504f & measurement <= 534f) method[step][1] = "+3";
                        else if (measurement > 534f & measurement <= 564f) method[step][1] = "-32";
                        else if (measurement > 564f & measurement <= 594f) method[step][1] = "-74";
                        else if (measurement > 594f & measurement <= 654f) method[step][1] = "-110";
                        df[step] = Integer.parseInt(method[step][1]);
                    }
                    break;
                }

                case 7: {
                    method[step][0] = "Минимальный диаметр диафиза (мм.)";
                    if (measurement <= 22f) method[step][1] = "+беск.";
                    else if (measurement > 36f) method[step][1] = "-беск";
                    else {
                        if (measurement > 22 & measurement <= 26) method[step][1] = "+78";
                        else if (measurement > 26f & measurement <= 28f) method[step][1] = "+45";
                        else if (measurement > 28f & measurement <= 30f) method[step][1] = "-5";
                        else if (measurement > 30f & measurement <= 32f) method[step][1] = "-45";
                        else if (measurement > 32f & measurement <= 36f) method[step][1] = "-124";
                        df[step] = Integer.parseInt(method[step][1]);
                    }
                    break;
                }

                case 8: {
                    method[step][0] = "Ширина проксимального эпифиза (мм.)";
                    if (measurement <= 84f) method[step][1] = "+беск.";
                    else if (measurement > 114f) method[step][1] = "-беск";
                    else {
                        if (measurement > 84 & measurement <= 90) method[step][1] = "+106";
                        else if (measurement > 90f & measurement <= 96f) method[step][1] = "+74";
                        else if (measurement > 96f & measurement <= 103f) method[step][1] = "0";
                        else if (measurement > 103f & measurement <= 108f) method[step][1] = "-100";
                        else if (measurement > 108f & measurement <= 114f) method[step][1] = "-134";
                        df[step] = Integer.parseInt(method[step][1]);
                    }
                    break;
                }
            }
        }
        //else -> сообщение об ошибке!
        return "ДК = "+ method[step][1];
    }


        //ОПРЕДЕЛЕНИЕ РОСТА:

//    public float getHeight_Pearson (float measurement, boolean man, boolean dry){
//    //определение прижизненного роста (по методике Pearson); категория низкого
//    //роста (по условной рубрикации Мартина).
//        float hp =0f;
//        if (man)//если пол исследуемой кости мужской
//        {
//            if (dry){ //если сотояние кости -> сухая
//
//                hp = 81.231f + 1.88f * (0.1f/*перевод в см*/ * measurement + 0.32f) - 1.26f;
//            }
//            else {//если состояние кости -> влажная
//
//                hp = 81.306f +1.88f * (0.1f/*перевод в см*/ * measurement +0.32f) - 1.26f;
//            }
//            //проверка вычисленного роста на соответствие методики.
//            if (hp <= 163.9f) hp = -hp;//методика не подходит -> изменение знака значения на "-"
//        }
//        else //если пол исследуемой кости женский
//        {
//            if(dry) {//если сотояние кости -> сухая
//
//                hp = 73.163f + 1.945f * (0.1f/*перевод в см*/ * measurement +0.33f) - 2f;
//            }
//            else { //если состояние кости -> влажная
//
//                hp = 72.844f + 1.945f * (0.1f/*перевод в см*/ * measurement +0.33f) - 2f;
//            }
//            //проверка вычисленного роста на соответствие методики.
//            if (hp<=152.9f) {
//                hp=-hp;
//            }
//        return hp;
//    }


//            public float getHeight_Telcca(float measurement, boolean man){
//        //определение прижизненного роста индивидуума по таблицам Telcca
//        if (man)//если исследуемая кость мужского пола:
//        {
//            if (measurement>=444f & measurement<=458f);/*проверка правильности выбора методики
//            по длине кости)*/
//            {
//                if (measurement==444f) height=165f;//с учетом "-2" для живого человека
//                else if (measurement>444f & measurement<448f) height=165f+(measurement-444f)/4f;
//                else if (measurement==448f)height=166f;//с учетом "-2" для живого человека
//                else if (measurement>448f & measurement<453f)height=166f+(measurement-448f)/5f;/*с учетом "-2" для живого человека*/
//                else if (measurement==453f)height=167f;//с учетом "-2" для живого человека
//                else if (measurement>453f & measurement<458f)height=167f+(measurement-453f)/5f;
//                else if (measurement==458f)height=168f;//с учетом "-2" для живого человека
//            }
//            //проверка правильности выбранной методики (с учетом вычисленного роста).
//            if (height>169.9f)//крайнее значение до категории высокого роста у мужчин
//                height=-height;//методика не подходит -> изменение знака значения на "-"
//
//        }
//        else //если исследуемая кость женского пола:
//        {
//            if (measurement>=414f & measurement<=430f) //проверка правильности выбора методики
//            {
//                if (measurement==414f) height=154f;//с учетом "-2" для живого человека
//                else if (measurement>414f & measurement<419f) height=154f+(measurement-414f)/5f;
//                else if (measurement==419f)height=155f;//с учетом "-2" для живого человека
//                else if (measurement>419f&measurement<425)height=155f+(measurement-419f)/6f;
//                else if (measurement==425f)height=156f;//с учетом "-2" для живого человека
//                else if (measurement>425f & measurement<430f)height=156f+(measurement-425f)/5f;
//                else if (measurement==430f)height=157f;//с учетом "-2" для живого человека
//            }
//            //проверка правильности выбранной методики (с учетом вычисленного роста).
//            if (height>158.9f)//крайнее значение до категории высокого роста у женщин
//                height=-height;
//        }
//        return height;
//    }

//    public float getHeight_Dupertuis_Hedden(float measurement, boolean man, boolean dry){
//
//        return height;
//    }
//    public  float getManouvrier(boolean man, boolean dry, float length){
//
//        return height;
//    }


}
