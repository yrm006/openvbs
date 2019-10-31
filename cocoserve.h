#ifndef _cocoserve_h_
#define _cocoserve_h_

#ifdef _WIN32
#else
 #include <netinet/in.h>
#endif



struct context{
    struct sockaddr_in  from;
    struct http         hdr;
    void*               nest;
};



// Implement on_xxx, on_ws_xxx, on_http_xxx functions for create your websocket server.
void on_start(struct sockaddr_in* bindto);
void on_stop();

//*NOTE* These functions MUST be thread safe.
int  on_ws_handshake(SOCKET sock, struct context* ctx);
int  on_ws_connect(SOCKET sock, struct context* ctx);
int  on_ws_data   (SOCKET sock, struct context* ctx, char* buf, size_t len, char opcode);
void on_ws_close  (SOCKET sock, struct context* ctx, char* buf, size_t len, char opcode);
int  on_http_request(SOCKET sock, struct context* ctx);



// Call these functions from each your on_start() and on_stop()
void ccs_start(struct sockaddr_in* bindto);
void ccs_stop();



#endif //_cocoserve_h_
