package bodewig.tools.postgresql;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;

import charlotte.tools.CsvData;
import charlotte.tools.FileTools;
import charlotte.tools.HugeQueue;
import charlotte.tools.StringTools;

/**
 * VM args: -Xmx1024M
 *
 * libs: postgresql-9.4.1212.jre6.jar
 *
 */
public class PgsqlExport {
	private String _host;
	private int _portNo;
	private String _dbName;
	private String _user;
	private String _password;

	public PgsqlExport(String host, int portNo, String dbName, String user, String password) {
		_host = host;
		_portNo = portNo;
		_dbName = dbName;
		_user = user;
		_password = password;
	}

	public String destRootDir = "C:/temp/PgsqlExport";

	private static int LIMIT_ROWCNT = 300;

	public void perform() throws Exception {
		HugeQueue dest = new HugeQueue();
		try {
			doSelect(dest, "SELECT schemaname, tablename FROM pg_tables");

			dest.poll(); // colcnt
			dest.poll(); // 1
			dest.poll(); // 2

			List<TableInfo> tableInfos = new ArrayList<TableInfo>();

			while(1 <= dest.size()) {
				TableInfo ti = new TableInfo();

				ti.scheme = dest.pollString();
				ti.name = dest.pollString();

				tableInfos.add(ti);
			}
			for(TableInfo ti : tableInfos) {
				doSelect(dest, "SELECT COUNT(*) FROM \"" + ti.scheme + "\".\"" + ti.name + "\"");

				dest.poll(); // colcnt
				dest.poll(); // 1

				ti.rowcnt = Integer.parseInt(dest.pollString());

				doSelect(dest, "SELECT * FROM \"" + ti.scheme + "\".\"" + ti.name + "\" LIMIT 1");

				ti.columns = new String[Integer.parseInt(dest.pollString())];

				for(int colidx = 0; colidx < ti.columns.length; colidx++) {
					ti.columns[colidx] = dest.pollString();
				}

				System.out.println(ti.scheme + "." + ti.name + ": " + ti.rowcnt + " rows, " + ti.columns.length + " columns");

				for(String column : ti.columns) {
					System.out.println("column: " + column);
				}
			}
			FileTools.rm(destRootDir);
			FileTools.mkdirs(destRootDir);

			for(TableInfo ti : tableInfos) {
				String wFile = FileTools.combine(destRootDir, ti.scheme + "." + ti.name + ".csv");

				CsvData.Stream csvStrm = new CsvData.Stream(wFile);
				csvStrm.writeOpen();
				try {
					for(int colidx = 0; colidx < ti.columns.length; colidx++) {
						csvStrm.add(ti.columns[colidx]);
					}
					csvStrm.endRow();

					for(int rowidx = 0; rowidx < ti.rowcnt; ) {
						int limit = Math.min(LIMIT_ROWCNT, ti.rowcnt - rowidx);

						doSelect(
								dest,
								"SELECT \"" +
								StringTools.join("\", \"", ti.columns) +
								"\" FROM \"" +
								ti.scheme +
								"\".\"" +
								ti.name +
								"\" ORDER BY \"ctid\" LIMIT " +
								limit +
								" OFFSET " +
								rowidx
								);

						dest.poll(); // colcnt

						for(int colidx = 0; colidx < ti.columns.length; colidx++) {
							dest.poll();
						}
						for(int rc = 0; rc < limit; rc++) {
							for(int colidx = 0; colidx < ti.columns.length; colidx++) {
								csvStrm.add(dest.pollString());
							}
							csvStrm.endRow();
						}
						rowidx += limit;
					}
				}
				finally {
					csvStrm.writeClose();
				}
			}
		}
		finally {
			FileTools.close(dest);
		}
	}

	private static class TableInfo {
		public String scheme;
		public String name;
		public int rowcnt;
		public String[] columns;
	}

	private void doSelect(HugeQueue dest, String sql) throws Exception {
		System.out.println("doSelect_sql: " + sql);

		dest.clear();

		Connection con = null;
		Statement stmt = null;
		ResultSet rs = null;

		try {
			Class.forName("org.postgresql.Driver");

			con = DriverManager.getConnection(
					"jdbc:postgresql://" + _host + ":" + _portNo + "/" + _dbName,
					_user,
					_password
					);
			stmt = con.createStatement();
			System.gc();
			rs = stmt.executeQuery(sql);
			System.gc();

			int colcnt = rs.getMetaData().getColumnCount();

			dest.add("" + colcnt);

			for(int colidx = 0; colidx < colcnt; colidx++) {
				dest.add(rs.getMetaData().getColumnName(colidx + 1));
			}
			while(rs.next()) {
				for(int colidx = 0; colidx < colcnt; colidx++) {
					Object oCell = rs.getObject(colidx + 1);
					String cell;

					if(oCell == null) {
						cell = "";
					}
					else if(oCell instanceof byte[]) {
						cell = StringTools.toHex((byte[])oCell);
					}
					else {
						cell = oCell.toString();
					}
					dest.add(cell);

					//System.gc();
				}
			}
			System.gc();
		}
		finally {
			FileTools.close(rs);
			FileTools.close(stmt);
			FileTools.close(con);
		}
	}
}
