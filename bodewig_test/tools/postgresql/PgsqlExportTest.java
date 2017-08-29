package bodewig_test.tools.postgresql;

import bodewig.tools.postgresql.PgsqlExport;

public class PgsqlExportTest {
	public static void main(String[] args) {
		try {
			main2();
		}
		catch(Throwable e) {
			e.printStackTrace();
		}
	}

	private static void main2() throws Exception {
		new PgsqlExport(
				"localhost", // host
				5432, // portNo
				"postgres", // dbName
				"postgres", // user
				"1111" // passphrase
				)
				.perform();
	}
}
