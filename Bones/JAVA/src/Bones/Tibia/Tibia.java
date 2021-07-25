package Bones.Tibia;

import Bones.Bones;

public class Tibia extends Bones {

    //FIELDS:



    //ARRAYS:
    //MANOUVRIER
    private static final  float[] mL_MANOUVRIER =  {
     //длина мужской кости.
            319f,324f,330f,335f,340f,346f,351f,357f,362f,368f,
            373f,378f,383f,389f,394f,400f,405f,410f,415f,420f
    };
    private static final  float[]wL_MANOUVRIER = {
   //длина женской кости.
            284f,289f,294f,299f,304f,309f,314f,319f,324f,329f,
            334f,340f,346f,352f,358f,364f,370f,376f,382f,388f
    };

    private static final float[]mGrPr_MANOUVRIER = {
    //рост мужчины.
            153.f,155.2f,157.1f,159.f,160.5f,162.5f,
            163.4f,164.4f,165.4f,166.6f,167.7f,168.6f,
            169.7f,171.6f,173.f,175.4f,176.7f,178.5f,
            181.2f,183.f
    };

    private static final float[]wGrPr_MANOUVRIER={
    //рост женщины.
            140.f,142.f,144.f,145.5f,147.f,148.8f,149.7f,
            151.3f,152.8f,154.3f,155.6f,156.8f,158.2f,
            159.5f,161.2f,163.f,165.f,167.f,169.2f,171.5f
    };

    //TROTTER_&_GLESER
    private float[]mGrPr_TrotterGleser;//рост мужчины.
    private float[]wGrPr_TrotterGleser;//рост женщины.
    private static final float [] mL_TROTTER_GLESER={
    //длина мужской большеберцовой кости (рост 152-198см):
            391f, 295f, 299f, 303f, 307f, 311f, 315f, 319f, 323f, 327f, 331f,
            335f, 339f, 343f, 347f, 351f, 355f, 359f, 363f, 367f, 371f, 375f,
            379f, 383f, 386f, 390f, 394f, 398f, 402f, 406f, 410f, 414f, 418f,
            422f, 426f, 430f, 434f, 438f, 442f, 446f, 450f, 454f, 458f, 462f,
            466f, 470f, 474f
    } ;

    /* по Алексееву (TrotterGleser)
               длина мужской большеберцовой кости (для мужчин европеоидов)
               (рост 152-198см):
               290f, 294f,298f,302f,306f,310f,314f,318f,323f,327f,331f,335f,
               339f, 343f,347f,352f,356f,360f,364f,368f,372f,376f,380f,385f,
               389f, 393f,397f,401f,405f,409f,414f,418f,422f,426f,430f,434f,
               438f, 442f,447f,451f,455f,459f,463f,467f,471f,475f,480f

               длина мужской большеберцовой кости (для мужчин монголоидов)
               (рост 152-198см):
               295f, 299f, 304f, 308f, 312f, 316f, 320f, 324f, 329f, 333f, 337f, 341f,
               345f, 350f, 354f, 358f, 362f, 366f, 371f, 375f, 379f, 383f, 387f, 391f,
               396f, 400f, 404f, 408f, 412f, 417f, 421f, 425f, 429f, 433f, 437f, 442f,
               446f, 450f, 454f, 458f, 463f, 467f, 471f, 475f, 479f, 483f, 488f


               длина муской большеберцовой кости (для мужчин негроидов)
               (рост 152-198см):
               304f, 309f, 313f, 319f, 323f, 328f, 333f, 337f, 342f, 346f, 351f, 355f,
               360f, 365f, 369f, 374f, 378f, 383f, 387f, 392f, 397f, 401f, 406f, 410f,
               415f, 419f, 424f, 428f, 433f, 438f, 442f, 447f, 451f, 456f, 460f, 465f,
               470f, 474f, 479f, 483f, 488f, 492f, 497f, 502f, 506f, 511f, 515f
   */

    private static final float[]wL_TROTTER_GLESER={
    //длина женской большеберцовой кости европеоиды
            // (рост 140-184см):
            271f,274f,277f,281f,284f,288f,291f,295f,298f,302f,305f,309f,
            312f,315f,319f,322f,326f,329f,333f,336f,340f,343f,346f,350f,
            353f,357f,360f,364f,367f,371f,374f,377f,381f,384f,388f,391f,
            395f,398f,402f,405f,409f,412f,415f,419f,422f
    };
    private static final float[]wnL_TROTTER_GLESER={
            //длина женской большеберцовой  кости (негроиды)
            // (рост 140-184см):
            375f, 279f, 283f, 287f, 291f, 295f, 299f, 303f, 308f, 312f, 316f,
            320f, 324f, 328f, 332f, 336f, 340f, 344f, 348f, 352f, 357f, 361f,
            365f, 369f, 373f, 377f, 381f, 385f, 389f, 393f, 397f, 401f, 406f,
            410f, 414f, 418f, 422f, 426f, 430f, 434f, 438f, 442f, 446f, 450f,
            454f
    };

    //TELCCA
    private float[]mGrPr_Telcca;//рост мужчины.
    private float[]wGrPr_Telcca;//рост женщины.
    private static final  float [] mL_TELCCA={
     //длина мужской кости.
            293f, 298f, 302f, 307f, 312f, 317f, 322f, 327f, 332f, 336f, 341f,
            346f, 350f, 355f, 360f, 365f, 370f, 375f, 379f, 384f, 389f, 394f,
            398f, 403f, 408f, 412f, 417f, 422f, 426f, 431f, 435f
    };
    private static final float [] wL_TELCCA = {
    //длина женской кости.
            268f, 274f, 280f, 285f, 290f, 295f, 300f, 306f, 311f, 316f, 321f,
            327f, 332f, 337f, 343f, 348f, 353f, 358f, 364f, 369f, 374f, 380f,
            385f, 390f, 395f, 400f, 405f, 411f, 416f, 421f, 426
    };


    //ENCAPSULATION:

    //CONSTRUCTOR
    Tibia (){
        super.upBound = 8;
        super.NAME1="большеберцовая кость";
        super.NAME2="большеберцовой кости";
        super.df = new int[upBound];
        super.method = new String[upBound][2];

            this.mGrPr_TrotterGleser = new float[47];
            //рост мужчины.
            for(int i = 0; i<48; i++){
               float tmp = 152f;
                mGrPr_TrotterGleser[i] = tmp++;
            }

            this.wGrPr_TrotterGleser = new float[45];
            //рост женщины.
            for(int i = 0; i<45; i++){
               float tmp = 140;
                wGrPr_TrotterGleser[i] = tmp++;
            }

            this.mGrPr_Telcca = new float[15];
            for(int i = 0; i<15; i++){
                float tmp = 155;
                mGrPr_Telcca [i] = tmp++;
            }

            this.wGrPr_Telcca = new float[15];
            for(int i = 0; i<15; i++){
                float tmp = 145;
                wGrPr_Telcca[i] = tmp++;
            }

    }

  //ОПРЕДЕЛЕНИЕ ПОЛА:
  @Override
  public String getDF(float measurement, int step) {
      //определение диагностических коэффициентов
      method[step][1]="НПВ";
      if (measurement>0f) {
          switch (step) {
              case 0: {
                  method[step][0] = "Общая длина кости (в естественном положении) (мм.)";
                  if (measurement <= 310f) method[step][1] = "+беск.";
                  else if (measurement > 410f) method[step][1] = "-беск";
                  else {
                      if (measurement > 310f & measurement <= 320f) method[step][1] = "+109";
                      else if (measurement > 320f & measurement <= 340f) method[step][1] = "+101";
                      else if (measurement > 340f & measurement <= 350f) method[step][1] = "+29";
                      else if (measurement > 350f & measurement <= 360f) method[step][1] = "-10";
                      else if (measurement > 360f & measurement <= 380f) method[step][1] = "-56";
                      else if (measurement > 380f & measurement <= 390f) method[step][1] = "-110";
                      else if (measurement > 390f & measurement <= 410f) method[step][1] = "-123";
                      df[step] = Integer.parseInt(method[step][1]);
                  }
               break;}
               
              case 1:{
                 method[step][0] ="Суставная длина (мм.)";
                  if (measurement <=290f)method[step][1] = "-беск.";
                  else if (measurement>390f)method[step][1] = "+беск.";
                  else {
                      if (measurement>290f & measurement<=310f)method[step][1]="+106";
                      else if (measurement>310f & measurement<=330f)method[step][1]="+70";
                      else if (measurement>330f & measurement<=350f)method[step][1]="-25";
                      else if (measurement>350f & measurement<=370f)method[step][1]="-96";
                      else if (measurement>370f & measurement<=390f)method[step][1]="-128";
                      df[step] = Integer.parseInt(method[step][1]);
                  }
              break;}
              
              case 2:{
                  method[step][0]="Ширина проксимального эпифиза (мм.)";
                  if (measurement<=69f)method[step][1]="-беск.";
                  else if (measurement>85f)method[step][1]="+беск.";
                  else {
                      if (measurement>69f & measurement<=71f)method[step][1]="+162";
                      else if (measurement>71f & measurement<=74f)method[step][1]="+93";
                      else if (measurement>74f & measurement<=76f)method[step][1]="+49";
                      else if (measurement>76f & measurement<=78f)method[step][1]="-58";
                      else if (measurement>78f & measurement<=80f)method[step][1]="-139";
                      else if (measurement>80f & measurement<=85f)method[step][1]="-159";
                      df[step] = Integer.parseInt(method[step][1]);
                  }
              break;}
              
              case 3:{
                  method[step][0]="Ширина дистального эпифиза (мм.)";
                  if (measurement<=46f) method[step][1] = "+беск.";
                  else if (measurement>59f)method[step][1]="-беск.";
                  else {
                      if (measurement > 46f & measurement <= 49f) method[step][1] = "+142";
                      else if (measurement > 49f & measurement <= 51f) method[step][1] = "+132";
                      else if (measurement > 51f & measurement <= 53f) method[step][1] = "+60";
                      else if (measurement > 53f & measurement <= 55f) method[step][1] = "-43";
                      else if (measurement > 55f & measurement <= 56f) method[step][1] = "-84";
                      else if (measurement > 56f & measurement <= 59f) method[step][1] = "-157";
                      df[3] = Integer.parseInt(method[step][1]);
                  }
              break;}
              
              case 4:{
                  method[step][0]="Сагитальный диаметр внешнего мыщелка (мм.)";
                  if (measurement<=40f) method[step][1]="+беск.";
                  else if (measurement>48f)method[step][1]="-беск.";
                  else {
                      if (measurement>40f & measurement<=42f)method[step][1]="+138";
                      else if (measurement>42f & measurement<=43f)method[step][1]="+61";
                      else if (measurement>43f & measurement<=44f)method[step][1]="-41";
                      else if (measurement>44f & measurement<=46f)method[step][1]="-89";
                      else if (measurement>46f & measurement<=48f)method[step][1]="-146";
                      df[step] = Integer.parseInt(method[step][1]);
                  }
              break;}
              
              case 5:{
                  method[step][0]="Окружность диафиза на уровне питательного отверстия (мм.)";
                  if (measurement<=79f)method[step][1]="+беск.";
                  else if (measurement>105f)method[step][1]="-беск.";
                  else {
                      if (measurement>79f & measurement<=83f)method[step][1]="+74";
                      else if (measurement>83f & measurement<=89f)method[step][1]="+43";
                      else if (measurement>89f & measurement<=93f)method[step][1]="-18";
                      else if (measurement>93f & measurement<=97f)method[step][1]="-57";
                      else if (measurement>97f & measurement<=105f)method[step][1]="-122";
                      df[step] = Integer.parseInt(method[step][1]);
                  }
              break;}
              
              case 6:{
                  method[step][0]="Диаметр середины диафиза (мм.)";
                  if (measurement<=18f)method[step][1]="+беск.";
                  else if (measurement>30f)method[step][1]="-беск.";
                  else {
                      if (measurement>18f & measurement<=20f)method[step][1]="+135";
                      else if (measurement>20f & measurement<=22f)method[step][1]="+72";
                      else if (measurement>22f & measurement<=25f)method[step][1]="-15";
                      else if (measurement>25f & measurement<=27f)method[step][1]="-69";
                      else if (measurement>27f & measurement<=30f)method[step][1]="-123";
                      df[step] = Integer.parseInt( method[step][1]);
                  }
              break;}
                  
              case 7:{
                  method[step][0] = "Площадь поперечного распила середины диафиза (мм.)";
                  if (measurement <= 280f) method[step][1] = "+беск.";
                  else if (measurement > 680f) method[step][1] = "-беск.";
                  else {
                      if (measurement > 281f & measurement <= 380f) method[step][1] = "+110";
                      else if (measurement > 380f & measurement <= 430f) method[step][1] = "+63";
                      else if (measurement > 430f & measurement <= 480f) method[step][1] = "-+37";
                      else if (measurement > 480f & measurement <= 530f) method[step][1] = "-68";
                      else if (measurement > 530f & measurement <= 680f) method[step][1] = "-125";
                      df[step] = Integer.parseInt(method[step][1]);
                  }
              }
          }
       }
      //else -> сообщение об ошибке!
        return "ДК = "+method[step][1];
  }

    //ОПРЕДЕЛЕНИЕ РОСТА:

    public float getHeight_Pearson (float  measurement_Pirson, boolean man, boolean dry) {
        /*определение прижизненного роста (по методике Pearson); категория низкого
         * роста (по условной рубрикации Мартина).*/
        if (man)//если пол исследуемой кости мужской
        {
            if (dry) //если сотояние кости -> сухая
            {
                height = 78.664f + 2.376f * (0.1f/*перевод в см*/ * measurement_Pirson - 0.96f) - 2f;
            } else //если состояние кости -> влажная
            {
                height = 78.807f + 2.376f * (0.1f/*перевод в см*/ * measurement_Pirson - 0.96f) - 2f;
            }
            //проверка вычисленного роста на соответствие методики.
            if (height <= 163.9) {height = -height;}//методика не подходит -> изменение знака значения на "-"
        } else //если пол исследуемой кости женский
        {
            if (dry) //если сотояние кости -> сухая
            {
                height = 74.774f + 2.352f * (0.1f/*перевод в см*/ * measurement_Pirson - 0.87f) - 1.26f;
            }
            else  //если состояние кости -> влажная
            {
                 height = 75.369f + 2.352f * (0.1f/*перевод в см*/ * measurement_Pirson - 0.87f) - 1.26f;
            }
            //проверка вычисленного роста на соответствие методики.
            if (height <= 152.9f) {height = -height;}

        }
        return height;
    }


//    public float getHeight_Telcca(float  measurement, boolean man){
//        //определение прижизненного роста индивидуума по таблицам Telcca
//        //категория СРЕДНЕГО и НИЗКОГО РОСТА
//        if (man)//если исследуемая кость мужского пола:
//        {
//            if ( measurement>=350f &  measurement<=365f);/*проверка правильности выбора методики
//            по длине кости)*/
//            {
//                if ( measurement==350f) height=165f;//с учетом "-2" для живого человека
//                else if ( measurement>350f &  measurement<355f) height=165f+( measurement-350f)/5f;
//                else if ( measurement==355f)height=166f;//с учетом "-2" для живого человека
//                else if ( measurement>355f &  measurement<360f)height=166f+( measurement-355f)/5f;/*с учетом "-2" для живого человека*/
//                else if ( measurement==360f)height=167f;//с учетом "-2" для живого человека
//                else if ( measurement>360f &  measurement<365f)height=167f+( measurement-360f)/5f;
//                else if ( measurement==365f)height=168f;//с учетом "-2" для живого человека
//            }
//            //проверка правильности выбранной методики (с учетом вычисленного роста).
//            if (height>169.9f)//крайнее значение до категории высокого роста у мужчин
//                height=-height;//методика не подходит -> изменение знака значения на "-"
//        }
//        else //если исследуемая кость женского пола:
//        {
//            if ( measurement>=327f &  measurement<=343f) //проверка правильности выбора методики
//            {
//                if ( measurement==327f) height=165f;//с учетом "-2" для живого человека
//                else if ( measurement>327f &  measurement<332f) height=165f+( measurement-327f)/5f;
//                else if ( measurement==332f)height=166f;//с учетом "-2" для живого человека
//                else if ( measurement>332f& measurement<337f)height=166f+( measurement-332f)/5f;
//                else if ( measurement==337f)height=167f;//с учетом "-2" для живого человека
//                else if ( measurement>337f &  measurement<343f)height=167f+( measurement-337f)/6f;
//                else if ( measurement==343f)height=168f;//с учетом "-2" для живого человека
//            }
//            //проверка правильности выбранной методики (с учетом вычисленного роста).
//            if (height>158.9f)//крайнее значение до категории высокого роста у женщин
//                height=-height;
//        }
//        return height;
//    }
//
//    public float getHeight_Dupertuis_Hedden(float  measurement, boolean man, boolean dry){
//    /*определение прижизненного роста (по методике Dupertuis_Hedden (1951)).
//    категория высокого роста (по условной рубрикации Мартина). */
//        if (man) //если пол исследуемой кости мужской
//        {
//            height=92.766f+2.178f*(0.1f/*перевод в см*/* measurement-0.96f)-2.f;
//            //проверка правильности методики (по условной рубрикации Мартина)
//            if (height<170.f) height=-height;//методика не подходит -> изменение знака значения на "-"
//        }
//        else //если пол исследуемой кости женский
//        {
//            height=71.52f+2.635f*(0.1f/*перевод в см*/* measurement-0.87f)-1.26f;
//            if (height<160.f) height=-height;//методика не подходит -> изменение знака значения на "-"
//        }
//        return height;
//    }

//    public String getManouvrier(boolean man, boolean dry, float length){
//    //
//    if (dry) length = length+0.2f;//для сухой кости к ее длине добавляем поправку
//    float tmp = 0.f;//временная переменная.
//
//        return "Manouvrier - " + str;
//    }


  }

