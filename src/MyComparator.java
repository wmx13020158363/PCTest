import java.util.Comparator;

public class MyComparator implements Comparator {

	@Override
	public int compare(Object arg0, Object arg1) {

        if(Double.parseDouble(((Books)arg0).getGrade()) > Double.parseDouble(((Books)arg1).getGrade())) {
        	return 1;
        }
		return 0;
	}

}
