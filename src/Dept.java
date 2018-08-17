
public class Dept {
    private String CRISID;
    private String UUID;
    private String name;

    public Dept(String CRISID, String UUID, String name) {
        this.CRISID = CRISID;
        this.UUID = UUID;
        this.name = name;
    }

    public String getCRISID() {
        return CRISID;
    }

    public void setCRISID(String CRISID) {
        this.CRISID = CRISID;
    }

    public String getUUID() {
        return UUID;
    }

    public void setUUID(String UUID) {
        this.UUID = UUID;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

}

enum DEPT_Field{
    CRISID (0),
    UUID (1),
    SOURCEREF (2),
    SOURCEID (3),
    NAME (4);

    private int value;

    private DEPT_Field(int value) {
        this.value = value;
    }

    public int getValue() {
        return this.value;
    }
}
