import oracledb, { BindParameter } from 'oracledb'

oracledb.fetchAsString = [oracledb.CLOB]
oracledb.outFormat = oracledb.OUT_FORMAT_OBJECT
oracledb.initOracleClient()

class Database {
  // eslint-disable-next-line no-use-before-define
  private static instance: Database
  private connection: oracledb.Connection | null = null

  public static getInstance(): Database {
    if (!Database.instance) {
      Database.instance = new Database()
    }
    return Database.instance
  }

  public async connect(
    user: string,
    password: string,
  ): Promise<oracledb.Connection> {
    if (this.connection) {
      return this.connection
    }

    try {
      const connect = await oracledb.getConnection({
        user,
        password,
        connectString: 'exacc01-scan.pbh:1521/sirh.pbh',
      })
      this.connection = connect
    } catch (error) {
      console.error('Erro ao conectar ao banco de dados Oracle:', error)
      throw error
    }

    return this.connection
  }

  public async disconnect(): Promise<void> {
    if (this.connection) {
      try {
        await this.connection.close()
        this.connection = null
      } catch (error) {
        console.error('Erro ao fechar a conexão com o banco de dados:', error)
      }
    }
  }

  public async executeQuery<T>(
    query: string,
    binds: { [key: string]: BindParameter },
  ): Promise<T[]> {
    if (!this.connection) {
      throw new Error('Conexão com o banco de dados não estabelecida.')
    }
    const result = await this.connection.execute<T>(query, binds)
    console.log(result)
    if (!result.rows) {
      return []
    }

    return result.rows
  }

  public async handleEvents(
    sql: string,
    binds: { [key: string]: BindParameter },
  ): Promise<void> {
    if (!this.connection) {
      throw new Error('Conexão com o banco de dados não estabelecida.')
    }

    await this.connection.subscribe('teste', {
      callback: () => {},
      sql,
      clientInitiated: true,
      binds,
    })
  }
}

export default Database
