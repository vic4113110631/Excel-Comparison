public class Researcher {
    String CRISID;
    String fullName;
    String translatedName;

    public Researcher(String CRISID, String fullName, String translatedName) {
        this.CRISID = CRISID;
        this.fullName = fullName;
        this.translatedName = translatedName;
    }

    public String getCRISID() {
        return CRISID;
    }

    public String getFullName() {
        return fullName;
    }

    public String getTranslatedName() {
        return translatedName;
    }

}
