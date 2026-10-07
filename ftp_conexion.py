"""Conexión FTP al hosting (produccion.millalemu.com), SIEMPRE con TLS explícito.

El 2026-10-07 por la tarde el hosting empezó a rechazar el FTP en claro:
"421-Sorry, cleartext sessions and weak ciphers are not accepted on this server".
La corrida de las 15:40 todavía subió; la de las 16:25 ya no, y quedó VERDE igual.

Los servidores que exigen TLS suelen exigir también REUTILIZAR la sesión TLS del canal de
control en el canal de datos (si no, "522 SSL connection failed: session reuse required");
ftplib no lo hace solo, por eso la subclase.
"""
import ftplib
import ssl


class _FTP_TLS_Reuso(ftplib.FTP_TLS):
    def ntransfercmd(self, cmd, rest=None):
        conn, size = ftplib.FTP.ntransfercmd(self, cmd, rest)
        if self._prot_p:
            conn = self.context.wrap_socket(conn, server_hostname=self.host,
                                            session=self.sock.session)
        return conn, size


def conectar(host, user, password, timeout=30):
    # Sin verificar el certificado: el hosting compartido presenta el del servidor
    # (*.bluehosting), no el de produccion.millalemu.com. Lo que se exige es el cifrado.
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    ftp = _FTP_TLS_Reuso(host, timeout=timeout, context=ctx)
    ftp.login(user, password)      # FTP_TLS.login hace AUTH TLS antes de mandar la clave
    ftp.prot_p()                   # canal de datos también cifrado
    return ftp
