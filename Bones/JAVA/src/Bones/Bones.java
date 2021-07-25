package Bones;

/*
Справочник:
условная рубрикация длины тела (по данным Мартина):
  -----------------------------------------------
    Категории   |    мужчины     |   женщины    |
  -----------------------------------------------
 1)Малая:       |               |               |
   карликовая   | до 129,9см    | до 120,9см    |
   очень малая  | 130,0-149,9см | 121,0-139,9см |
   малая        | 150,0-159,9см | 140,0-148,9см |
  -----------------------------------------------
 2)Средняя      |               |               |
   ниже среднего| 160,0-163,9см | 149,0-152,9см |
   средняя      | 164,0-166,9см | 153,0-155,9см |
   выше среднего| 167,0-169,9см | 156,0-158,9см |
  -----------------------------------------------
 3)Большая      |               |               |
   большая      | 170,0-179,9см | 159,0-167,9см |
   очень большая| 180,0-199,0см | 168,0-186,9см |
   гигантская   | более 200,0см | более 187,0см |
  -----------------------------------------------
*/

public abstract class Bones {
    //родительский класс

    //FIELDS:
    private boolean man = true; //определяемый пол:   true (man); false (woman)
    private boolean dry = true;//состояние кости: сухая = true; влажная = false;
    public int sumDF = 0;//сумма диагностических коэффициентов.

    public int df[];// массив значений диагностических коэффициентов.

    public float height = 0f;//вычисленный рост индивидуума.
    public String sumDf = "Пол не определен.";//строковое представление суммы диагностических коэффициентов.
    public String sex = "Половая принадлежность костных останков не определена";//пол индивидуума.
    public String NAME1 = "плечевая кость";
    public String NAME2 = "плечевой кости";
    public int upBound; //например=8 -> максимальное число процедур определения ДК (счетчик)

    public String[][] method;  //для подстановки в массив в качестве индекса.
    // массив значений   [0] [0] - название метода;  [0] [1] - значение ДК

    public float[] mGrPr_Manouvrier = {
            //рост мужчины  (153-183см).
            153f, 155.2f, 157.1f, 159f, 160.5f, 162.5f, 163.4f, 164.4f, 166.6f,
            167.7f, 168.6f, 169.7f, 171.6f, 173f, 175.4f, 176.7f, 178.5f,181.2f,
            183f
    };

    public float[]wGrPr_Manouvrier = {
            //рост женщины (140-171,5см).
            140f, 142f, 144f, 145.5f, 147f, 148.8f, 149.7f, 151.3f, 152.8f,
            154.3f, 155.6f, 156.8f, 158.2f, 159.5f, 161.2f, 163f, 165f, 167f,
            169.2f, 171.5f
    };



    //ENCAPSULATION:
    private void setMan(boolean man) {
        this.man = man;
    }

    public boolean isMan() {
        return man;
    }

    private void setDry(boolean dry) {
        this.dry = dry;
    }

    public boolean isDry() {
        return dry;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }

    public String getSex() {
        return sex;
    }

    public void setHeight(float height) {
        this.height = height;
    }

    public float getHeight() {
        return height;
    }


    //FINAL:
    private static final String EXAMINATION = "методика расчета прижизненного роста индивидуума";
    private static final String HEIGHT = "Прижизненный рост индивидуума";
    private static final String METHOD = ", вычисленный по методике ";
    private static final String DUPERTUIS_HEDDEN = "Dupertuis_Hedden ";
    private static final String TELCCA = "Telcca ";
    private static final String MANOUVRIER = "Manouvrier ";
    private static final String TROTTER_GLESER = "Trotter & Gleser ";
    private static final String PEARSON = "Pearson ";
    private static final String NO_EXAM = "не подходит";
    private static final String WOM_GENDER = "женского пола";
    private static final String MAN_GENDER = "мужского пола";
    private static final String DEFINABLE_GENDER = "Определяемый пол ";
    private static final String RELIABLY_MAN = "достоверно мужской.";
    private static final String RELIABLY_WOM = "достоверно женский.";
    private static final String PROBABLY_MAN = ", вероятно, мужской.";
    private static final String PROBABLY_WOM = ", вероятно, женский.";
    private static final String MAN_MIN = "Рост мужчины менее 130см!";
    private static final String MAN_MAX = "Рост мужчины более 200см!";
    private static final String WOM_MIN = "Рост женщины менее 121см!";
    private static final String WOM_MAX = "Рост женщины более 187см!";
    private static final String ATTANTION_MIN = "Категория роста -> карликовая, или кость может принадлежать ребенку!";
    private static final String ATTANTION_MAX = "Категория роста -> гигантская. Проверте корректность введенной длины кости!";

    //ABSTRACT
    public abstract String getDF(float measurement, int step);
//    public abstract float getHeight_Pearson (float maxLength, boolean man, boolean dry);
//    public abstract float getHeight_Telcca(float maxLength, boolean man);
//    public abstract float getHeight_Dupertuis_Hedden(float maxLength, boolean man, boolean dry);
//    public abstract float getManouvrier(float maxLength, boolean man, boolean dry);
//    public abstract float gerTrotterGleser (float maxLength, boolean man, boolean dry);


    public String getSumDF(String method[][], int df[]) {
        //подсчет суммы диагностических коэффициентов.
//  1).циклическая обработка массивов:
//      посчет суммы числовых коэффициентов
        int sum = 0;
        for (int x = 0; x > upBound; x++) {
            sum = +df[x];
        }
//      подсчет количества значений "-беск.", "+беск." и "НПВ"
        int w = 0;  //для подсчета "+беск."
        int m = 0;  //для подсчта "-беск."
        for (int i = 0; i > upBound; i++) {
            if (method[i][1].equalsIgnoreCase("+беск."))//если df - это "+беск."
                ++w;
            else if (method[i][1].equalsIgnoreCase("-беск."))//если df - это "-беск.
                ++m;
        }
//  2).определение итогового значения суммы коэффициентов:
        if (w == 0 & m == 0)
            this.sumDf = "" + sum;
        else {
            if (w > m & w > 0) this.sumDf = "+беск.";
            else if (m > w & m > 0) this.sumDf = "-беск.";
        }
        return this.sumDf;
    }

    public String sumDF_toString(String sumDF) {
//    строковое представление суммы диагностичесих коэфициентов
        String toString = "Половая принадлежность исследуемой "; //+ NAME2 +"не определена.";
        String mark = "+";

//        if (sumDF>0){
//            if(sumDF<=-128)toString="Сумма ДК = " + sumDf +". Исследуемая "+NAME1+" могла принадлежать "+
//                    "индивидууму мужского пола.";
//            else if(sumDF>-128 & sumDF<0)toString="Сумма ДК = " + sumDf +". Исследуемая "+NAME1+" могла принадлежать "+
//                    "индивидууму, возможно, мужского пола.";
//            else if(sumDF>0 & sumDF<128)toString="Сумма ДК = " + mark+sumDf +". Исследуемая "+NAME1+" могла принадлежать "+
//                    "индивидууму, возможно, женского пола.";
//            else if(sumDF>=128)toString="Сумма ДК = " +mark+sumDf +". Исследуемая "+NAME1+" могла принадлежать "+
//                    "индивидууму женского пола.";
//        }
        return toString;
    }

    public String heigthToString(String method, String sex, float flHeigth) {/*строка="Прижизненный рот индивидуума мужского/женского пола, вычисленный по методике
    (название методики) = ХХХ.Хсм"*/
        String height = "Методика ";//+ method + "для вычисления прижизенного роста индивидуума "+ Bones.DISAGREE;
//        if (flHeigth>0){
//            height= Bones.HEIGHT+sex+ Bones.METHOD+method+" = "+ String.format("% .2f",flHeigth + "см");
//        }
        return height;
    }
}

