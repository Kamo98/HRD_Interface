using Npgsql;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Windows.Forms;

namespace PersonnelDeptApp1
{
	class Connection
	{
		public enum Role
		{
			admin,			//Админ
			accounting,		//Бюро учёта
			reception,		//Бюро приёма
			unknown
		}

		private static Dictionary<string, Role> str2role = new Dictionary<string, Role>()
		{
			{"admin", Role.admin},
			{"accounting", Role.accounting},
			{"reception", Role.reception}
		};

		private static Dictionary<Role, string> role2str = new Dictionary<Role, string>()
		{
			{Role.admin, "admin"},
			{Role.accounting, "accounting"},
			{Role.reception, "reception"}
		};


		private static NpgsqlConnection npgSqlConnection = null;
		private static Connection connect = null;
		private static Role role = Role.unknown;

		private Connection (string login, string pass)
		{
			string connectionString = "Server = hrd.cx7kyl76gv42.us-east-2.rds.amazonaws.com; DataBase = HRD; Integrated Security = false; User Id = " + login + "; password = " + pass;

			//Создание соединения с БД
			npgSqlConnection = new NpgsqlConnection(connectionString);

			try
			{
				npgSqlConnection.Open();        //Открываем соединение

				//Получаем роль пользователя
				string strCom = "select \"role_name\" from \"information_schema\".\"applicable_roles\" where \"grantee\" = '" + login +"'";
				NpgsqlCommand command = new NpgsqlCommand(strCom, npgSqlConnection);
				NpgsqlDataReader reader = command.ExecuteReader();

				if (reader.HasRows)
				{
					string roleStr = "";
					foreach (DbDataRecord rec in reader)
					{
						roleStr = rec.GetString(0);
						break;
					}
					role = str2role[roleStr];

				} else
				{
					npgSqlConnection = null;
					MessageBox.Show("Подключение НЕ выполнено.\nРоль пользователя " + login + " не найдена\n");
				}
				reader.Close();

			}
			catch (NpgsqlException e)
			{
				npgSqlConnection = null;
				MessageBox.Show("Подключение НЕ выполнено.\nПроверьте правильность введённых логина и пароля\n" + e.Message + "\n");
			}
		}
		

		public static Role get_role()
		{
			return role;
		}

		public static string get_role_str()
		{
			return role2str[role];
		}

		public static NpgsqlConnection get_connect()
		{
			return npgSqlConnection;
		}


		public static bool create_instance(string login, string pass)
		{
			if (connect == null)
				connect = new Connection(login, pass);

			return (npgSqlConnection != null);			
		}
		

	
		public static void close_connection ()
		{
			if (npgSqlConnection != null)
			{
				npgSqlConnection.Close();
			}
			npgSqlConnection = null;
			connect = null;
			role = Role.unknown;
		}
		
	}
}
