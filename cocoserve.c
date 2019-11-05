#include <errno.h>
#include <stdio.h>
#include <stdint.h>

#ifdef _WIN32
 #include <winsock2.h>
 #include <windows.h>
 #include <process.h>
#else
 #include <arpa/inet.h>
 #include <sys/fcntl.h>
 #include <poll.h>
 #include <sys/socket.h>
 
 #include <unistd.h>
 #include <pthread.h>
#endif

#include "http.h"
#include "websocket.h"
#include "cocoserve.h"

#define LOG(m,...)        fprintf(stdout,m"\n",##__VA_ARGS__)
#define LO_(...)          
#define CHECK(e,m,...)    if(e)fprintf(stderr,m"\n",##__VA_ARGS__),exit(errno)



#ifdef _WIN32
    #define struct       
    #define poll         WSAPoll
    #define pollfd       WSAPOLLFD
    #define socklen_t    int
    #define usleep(m)    Sleep(m/1000)
    #define pthread_self GetCurrentThreadId
    
    void* main_poll(void* arg);
    UINT WINAPI start_routine_win32(void* arg){
        main_poll(arg); return 0;
    }
    
    typedef HANDLE  pthread_t;
    typedef void    pthread_attr_t;
    inline int pthread_create(pthread_t * thread, pthread_attr_t * attr, void* (*start_routine)(void *), void * arg){
        return (*thread = (HANDLE)_beginthreadex(NULL, 0, start_routine_win32, arg, 0, NULL)) ? 0 : -1;
    }
    inline int pthread_join(pthread_t th, void **thread_return){
        return (WaitForSingleObject(th, INFINITE) == WAIT_OBJECT_0) ? 0 : -1;
    }
#else
    #define CloseHandle(...)
#endif



#define BACKLOG         0

#ifdef _WIN32
    // Polling threads.
    #define POLLTRD         40

    // Polling sockets/thread.
    #define POLLUNT         (100000/POLLTRD)
#else
    // Polling threads.
    #define POLLTRD         40

    // Polling sockets/thread.
    #define POLLUNT         (10240/POLLTRD)                                         // OSX: sysctl kern.maxfilesperproc=10240
#endif

// Payload buffer length.
#define FRMEBUF         65536









typedef void (*PFUNC)();
typedef PFUNC (*ACTION)(SOCKET, struct context*);

static SOCKET g_lstn = INVALID_SOCKET;                                              // for accept connection from clients.
static SOCKET g_ctrl = INVALID_SOCKET;                                              // for reveive ctrl command from main thread.



static
PFUNC on_frame(SOCKET sock, struct context* ctx){
    ACTION next = on_frame;
    
    char aBuf[FRMEBUF];
    aBuf[sizeof(aBuf)-1] = '\0';                                                    // For safety.
    
    char opcode;
    int n = ws_recv(sock, aBuf, sizeof(aBuf)-1, 0, &opcode);
                                                                                    LO_("on_frame:%d, %x", n, opcode);
    
    if(opcode & WS_OP_CTRL){
        if(opcode == WS_OP_CLOSE){
            on_ws_close(sock, ctx, aBuf, n, opcode);
            next = NULL;
        }
    }else
    if(opcode == WS_OP_TEXT || opcode == WS_OP_BINARY){
        if(!on_ws_data(sock, ctx, aBuf, n, opcode)){
            on_ws_close(sock, ctx, aBuf, n, opcode);
            next = NULL;
        }
    }else{
        on_ws_close(sock, ctx, aBuf, n, opcode);
        next = NULL;
    }
    
    return (PFUNC)next;
}



static
PFUNC on_http(SOCKET sock, struct context* ctx){
                                                                                    LO_("on_http.");
    ACTION next = NULL;
    
    int n = 0;
    while( 
        (recv(sock, ctx->hdr.m_aBuf+n, 1, 0) == 1)  &&
        !(
            ctx->hdr.m_aBuf[n-3]=='\r' && 
            ctx->hdr.m_aBuf[n-2]=='\n' && 
            ctx->hdr.m_aBuf[n-1]=='\r' && 
            ctx->hdr.m_aBuf[n-0]=='\n' &&
        1)                                          &&
        (++n < HTTP_BUFSIZE)                        &&
    1);
    
    if(0 < http_parse(&ctx->hdr, n)){
        if(atoi(http_header(&ctx->hdr, "Sec-WebSocket-Version")) == WS_VER){
            if(on_ws_handshake(sock, ctx)){
                if(ws_handshake(sock, &ctx->hdr)){
                    if(on_ws_connect(sock, ctx)){
                        next = on_frame;
                    }
                }
            }else{
                // header will sent by 'on_ws_handshake'.
            }
        }else{
			if(on_http_request(sock, ctx)){
				next = on_http;
			}else{
                // header will sent by 'on_http_request'.
            }
        }
    }else{
                                                                                    LO_("!http_parse.");
		char aBuf[256];
		aBuf[sizeof(aBuf)-1] = '\0';                                                // For safety.
		
		int n = 0;
		n += snprintf(aBuf+n, sizeof(aBuf)-n,
			"HTTP/1.0 400 Bad Request\r\n"
			"\r\n"
		);
		
		send(sock, aBuf, n, 0);
    }
    
                                                                                    LO_("on_http:%p:%d:%s.", next, n, ctx->hdr->m_aBuf);
    return (PFUNC)next;
}



static
void* main_poll(void* arg){
                                                                                    LO_("thread started.");
    struct pollfd* fds = (struct pollfd*)malloc(sizeof(struct pollfd) * POLLUNT);
    {
        // set ctrl socket
        fds[0].fd = g_ctrl;
        fds[0].events = POLLIN;
        
        // set listen socket
        fds[1].fd = g_lstn;                                                         // fds[0].fd is copy of g_lstn.
        fds[1].events = POLLIN;
        
        int i;
        for(i=2; i<POLLUNT; ++i){
            fds[i].fd = INVALID_SOCKET;
        }
    }
    
    ACTION* acts = (ACTION*)malloc(sizeof(ACTION) * POLLUNT);
    struct context* ctxs = (struct context*)malloc(sizeof(struct context) * POLLUNT);
    
    int pls = 2;                                                                    // 0:g_ctrl, 1:g_lstn
    
    
    
    while(g_lstn != INVALID_SOCKET){
                                                                                    LO_("polling-%p...", pthread_self());
        int n = poll(fds, pls, -1);
                                                                                    LO_("poll-%p: n=%d, fds[0].revents=%x, fds[1].revents=%x.", pthread_self(), n, fds[0].revents, fds[1].revents);
        
        if(fds[0].revents){                                                         // ctrl event occured.
                                                                                    LO_("fds[0].revents=%x.", fds[0].revents);
            --n;
        }
        
        if(0<n && fds[1].revents){                                                  // listening socket's accept occured.
                                                                                    LO_("fds[1].revents=%x.", fds[1].revents);
            struct sockaddr_in afrom;
            socklen_t   nfrom = sizeof(afrom);
            
            SOCKET as = accept(fds[1].fd, (struct sockaddr*)&afrom, &nfrom);
                                                                                    LO_("accept: as=%d, from=%x:%d.", as, ntohl(afrom.sin_addr.s_addr), ntohs(afrom.sin_port));
            if(as != INVALID_SOCKET){
                                                                                    CHECK(POLLUNT <= pls, "!pls:%d", pls);
                {                                                                   // set to blocking mode.
                #ifdef _WIN32
                    u_long mode = 0;
                    ioctlsocket(as, FIONBIO, &mode);
                #else
                    fcntl(as, F_SETFL, fcntl(as, F_GETFL)&~O_NONBLOCK);
                #endif
                }
                
                fds[pls].fd = as;
                fds[pls].events = POLLIN;
                fds[pls].revents = 0;
                acts[pls] = on_http;
                ctxs[pls].from = afrom;
                http_reset(&ctxs[pls].hdr);
                ctxs[pls].nest = NULL;
                
                ++pls;
                
                if(POLLUNT <= pls){                                                 // if it is last polling slot.
                    fds[1].events = 0;                                              // set to 'no more accept' mode.
                }
            }
            --n;
        }
        
        int i;
        for(i=2; 0<n && i<pls; ++i){
            if(fds[i].revents){                                                     // recv occured.
                                                                                    LO_("fds[%d].revents=%x.", i, fds[i].revents);
                acts[i] = (ACTION)acts[i](fds[i].fd, &ctxs[i]);
                if(acts[i]){
                                                                                    LO_("acts[%d]:%s:%d,%d.", i, ctxs[i].hdr.m_pURI, (int)ctxs[i].hdr.m_nBuf, (int)ctxs[i].hdr.m_nParsed);
                }else{
                                                                                    LO_("!acts[%d]:%s:%d,%d.", i, ctxs[i].hdr.m_pURI, (int)ctxs[i].hdr.m_nBuf, (int)ctxs[i].hdr.m_nParsed);
                    CLOSE(fds[i].fd);
                                                                                    LO_("close: fds[%d].fd=%d.", i, fds[i].fd);
                    
                    --pls;
                    
                    {                                                               // move last to current for deflag.
                        fds[i].fd = fds[pls].fd;
                        fds[i].events = fds[pls].events;
                        acts[i] = acts[pls];
                        
                        fds[pls].fd = INVALID_SOCKET;
                        acts[pls] = NULL;
                    }
                    
                    if(fds[1].events == 0){                                         // if now is 'no more accept' mode.
                        fds[1].events = POLLIN;                                     // reset to accept mode.
                    }
                }
                --n;
            }
        }
                                                                                    LO_("poll_: n=%d, fds[0].revents=%x, fds[1].revents=%x.", n, fds[0].revents, fds[1].revents);
    }
                                                                                    LO_("pls1: %d", pls);
    
    
    
    int i;
    for(i=pls-1; 1<i; --i){
        CLOSE(fds[i].fd);
                                                                                    LO_("close: fds[%d].fd=%d.", i, fds[i].fd);
        --pls;
    }
                                                                                    LO_("pls2: %d", pls);
    
    free(ctxs);
    free(acts);
    free(fds);
    
                                                                                    LO_("thread stopped.");
    return arg;
}









static SOCKET    g_ctrler       = INVALID_SOCKET;
static pthread_t g_tid[POLLTRD] = {};



void ccs_start(struct sockaddr_in* bindto){
    int n;
    
    
    
    g_lstn = socket(PF_INET, SOCK_STREAM, 0);
                                                                                    CHECK(g_lstn == INVALID_SOCKET, "!socket:g_lstn");
    n = bind(g_lstn, (struct sockaddr*)bindto, sizeof(*bindto));
                                                                                    CHECK(n != 0, "!bind:%d", n);
    n = listen(g_lstn, BACKLOG);
                                                                                    CHECK(n == -1, "!listen:%d", n);
    
    
    
    struct sockaddr_in ctrlto = *bindto;
    if(ctrlto.sin_addr.s_addr == INADDR_ANY)
    {
        ctrlto.sin_addr.s_addr = 0x0100007f;                                        // localhost(127.0.0.1)
    }
    
    g_ctrler = socket(PF_INET, SOCK_STREAM, 0);                                     // for send ctrl command to polling threads.
                                                                                    CHECK(g_ctrler == INVALID_SOCKET, "!socket:g_ctrler");
    n = connect(g_ctrler, (struct sockaddr*)&ctrlto, sizeof(ctrlto));
                                                                                    CHECK(n != 0, "!connect:g_ctrler:%d", n);
    g_ctrl = accept(g_lstn, NULL, NULL);
                                                                                    CHECK(g_ctrl == INVALID_SOCKET, "!accept:g_ctrl");
    
    
    
    {
    #ifdef _WIN32
        u_long mode = 1;
        ioctlsocket(g_lstn, FIONBIO, &mode);
    #else
        fcntl(g_lstn, F_SETFL, fcntl(g_lstn, F_GETFL)|O_NONBLOCK);
    #endif
    }
    
    #ifdef SNDBUF
    {
        int nSndBuf = SNDBUF;
        setsockopt(g_lstn, SOL_SOCKET, SO_SNDBUF, &nSndBuf, sizeof(nSndBuf));
                                                                                    LO_("set SO_SNDBUF:%d.", nSndBuf);
    }
    #endif
    
    #ifdef __APPLE__
    {
        int nNoSigPipe = 1;
        setsockopt(g_lstn, SOL_SOCKET, SO_NOSIGPIPE, &nNoSigPipe, sizeof(nNoSigPipe));
                                                                                    LO_("set SO_NOSIGPIPE:%d.", nNoSigPipe);
    }
    #endif
    
    int i;
    for(i=0; i<POLLTRD; ++i){
        n = pthread_create(&g_tid[i], NULL, main_poll, NULL);
                                                                                    CHECK(n != 0, "!pthread_create.");
    }
    
                                                                                    LO_("server started.");
    
}



void ccs_stop(){
    SOCKET stopping = g_lstn;
    
    g_lstn = INVALID_SOCKET;
    send(g_ctrler, "\0", 1, 0);
    
    int i;
    for(i=0; i<POLLTRD; ++i){
        pthread_join(g_tid[i], NULL);
        CloseHandle(g_tid[i]);
    }
    
    CLOSE(g_ctrler);
    CLOSE(g_ctrl);
    CLOSE(stopping);
                                                                                    LO_("close.");
    
                                                                                    LO_("server stopped.");
    
}


