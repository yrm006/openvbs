#ifndef _websocket_h_
#define _websocket_h_

#include "http.h"
#include <stdlib.h>



#ifdef _WIN32
    #define CLOSE closesocket
#else
    typedef int SOCKET;
    #define INVALID_SOCKET (SOCKET)-1
    #define CLOSE close
#endif



#define WS_OP_TEXT      0x01
#define WS_OP_BINARY    0x02
#define WS_OP_CLOSE     0x08
#define WS_OP_PING      0x09
#define WS_OP_PONG      0x0a
#define WS_OP_CTRL      0x08
#define WS_MASK         0x80



// Handshake buffer length.
#define WS_HBUF 1024

// WebSocket Verson
#define WS_VER 13



int ws_upgrade  (SOCKET sock, const char* pURI, const char* pHost, const char* pHeaders, struct http* res);
int ws_handshake(SOCKET sock, struct http* req);

int ws_send(SOCKET sock, const char* buf, size_t len, int flags, char* mask, char  opcode);
int ws_recv(SOCKET sock,       char* buf, size_t len, int flags,             char* popcode);



#endif //_websocket_h_
