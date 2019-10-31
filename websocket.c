#ifdef _WIN32
 #include <winsock2.h>
#else
 #include <sys/socket.h>
#endif

#include "websocket.h"
#include <stdio.h>
#include <stdint.h>
#include <string.h>

#define WSLOG(m,...)        fprintf(stdout,m"\n",##__VA_ARGS__)
#define WSLO_(...)          



#ifdef _WIN32
    #define snprintf _snprintf
    
    #define LITTLE_ENDIAN  1234
    #define BIG_ENDIAN     4321
    #define BYTE_ORDER     LITTLE_ENDIAN
#else
#endif



static void sha1_calc(const void* src, const int bytelength, unsigned char* hash);
static void base64_encode_binary(char *out, const unsigned char *in, size_t len);









int ws_upgrade(SOCKET sock, const char* pURI, const char* pHost, const char* pHeaders, struct http* res){
    char aBuf[WS_HBUF];
    aBuf[sizeof(aBuf)-1] = '\0';                                                    // For safety.
    
    
    
    unsigned char keysource[16+1] = {
        0x5c, 0x76, 0x02, 0x55, 0xe2, 0x6e, 0xfb, 0x88,
        0x14, 0x87, 0x14, 0xa1, 0xf3, 0x80, 0x51, 0xde,
        0x00
    };
    
    char aKey[((16+2)/3)*4 + 1];
    base64_encode_binary(aKey, keysource, 16);
    
    
    
    int n = 0;
    n += snprintf(aBuf+n, sizeof(aBuf)-n,
        "GET %s HTTP/1.1\r\n"
        "Host: %s\r\n"
        "Upgrade: websocket\r\n"
        "Connection: Upgrade\r\n"
        "Sec-WebSocket-Key: %s\r\n"
        "Sec-WebSocket-Version: %d\r\n"
        "%s"
        "\r\n",
        pURI,
        pHost,
        aKey,
        WS_VER,
        pHeaders
    );
    
    send(sock, aBuf, n, 0);
    
    
    
    int m = recv(sock, res->m_aBuf, HTTP_BUFSIZE, 0);
    if(m <= 0) return m;
    
    if(
       0 < http_parse(res, m)                                                                   &&
       atoi(res->m_pCode) == 101                                                                &&
       strcmp(http_header(res, "Sec-WebSocket-Accept"), "iHmS7O9QVLcIbhWI7vqvv2y79m4=") == 0    )
    {
        return 1;
    }
    
    return -1;
}



int ws_handshake(SOCKET sock, struct http* req){
//    if(atoi(http_header(req, "Sec-WebSocket-Version")) == WS_VER){
        char aBufAppended[((16+2)/3)*4 + 36 + 1];
        int nBufAppended = snprintf(aBufAppended, sizeof(aBufAppended),
            "%s" "258EAFA5-E914-47DA-95CA-C5AB0DC85B11",
            http_header(req, "Sec-WebSocket-Key")
        );
        
        unsigned char hash[20+1];
        sha1_calc(aBufAppended, nBufAppended, hash);
        
        hash[20] = 0;                                                               // For base64_encode_binary SPEC...
        char aAccept[((20+2)/3)*4 + 1];
        base64_encode_binary(aAccept, hash, 20);
        
        char aBuf[WS_HBUF];
        aBuf[sizeof(aBuf)-1] = '\0';                                                // For safety.
        
        int n = 0;
        n += snprintf(aBuf+n, sizeof(aBuf)-n,
            "HTTP/1.1 101 Switching Protocols\r\n"
            "Connection: Upgrade\r\n"
            "Upgrade: websocket\r\n"
            "Sec-WebSocket-Accept: %s\r\n"
            "\r\n",
            aAccept
        );
        
        return send(sock, aBuf, n, 0);
//    }

//    return 0;
}



int ws_send(SOCKET sock, const char* buf, size_t len, int flags, char* mask, char opcode){
    int n;
    
    char aHeader[14];
    
    aHeader[0] = 0;
    aHeader[0] |= 0x80;//fin :0x80
    aHeader[0] |= 0x00;//rsv1:0x40
    aHeader[0] |= 0x00;//rsv2:0x20
    aHeader[0] |= 0x00;//rsv3:0x10
    aHeader[0] |= opcode;
    
    aHeader[1] = 0;
    aHeader[1] |= mask ? WS_MASK : 0;
    
    if(0xFFFF < len){
        aHeader[1] |= 0x7f;
        #if BYTE_ORDER == BIG_ENDIAN
            *(uint64_t*)&aHeader[2] = len;
        #else
            aHeader[2+0]=((char*)&len)[7];
            aHeader[2+1]=((char*)&len)[6];
            aHeader[2+2]=((char*)&len)[5];
            aHeader[2+3]=((char*)&len)[4];
            aHeader[2+4]=((char*)&len)[3];
            aHeader[2+5]=((char*)&len)[2];
            aHeader[2+6]=((char*)&len)[1];
            aHeader[2+7]=((char*)&len)[0];
        #endif
        
        if(mask){
            *(uint32_t*)&aHeader[10] = *(uint32_t*)mask;
            n = send(sock, aHeader, 2+8+4, flags);
        }else{
            n = send(sock, aHeader, 2+8+0, flags);
        }
    }else
    if(0x7d < len){
        aHeader[1] |= 0x7e;
        #if BYTE_ORDER == BIG_ENDIAN
            *(uint16_t*)&aHeader[2] = len;
        #else
            aHeader[2+0]=((char*)&len)[1];
            aHeader[2+1]=((char*)&len)[0];
        #endif
        
        if(mask){
            *(uint32_t*)&aHeader[4] = *(uint32_t*)mask;
            n = send(sock, aHeader, 2+2+4, flags);
        }else{
            n = send(sock, aHeader, 2+2+0, flags);
        }
    }else{
        aHeader[1] |= len;
        
        if(mask){
            *(uint32_t*)&aHeader[2] = *(uint32_t*)mask;
            n = send(sock, aHeader, 2+0+4, flags);
        }else{
            n = send(sock, aHeader, 2+0+0, flags);
        }
    }
    
    if(n <= 0){
        return n;
    }
    
    if(mask){
        char* tmp = (char*)malloc(len);
        
        size_t i;
        for(i=0; i<len; ++i){
            tmp[i] = buf[i] ^ mask[i % 4];
        }
        
        n = send(sock, tmp, len, flags);
        
        free(tmp);
    }else{
        n = send(sock, buf, len, flags);
    }
    
    return n;
}



int ws_recv(SOCKET sock, char* buf, size_t len, int flags, char* popcode){
    (void)flags;
    int n;
    
    if(popcode)
        *popcode = 0;
    
    char aHeader[14];
    if((n = recv(sock, aHeader, 2, 0)) <= 0) return n;
    
    char fin = aHeader[0] & 0x80;
    char rsv1 = aHeader[0] & 0x40;  (void)rsv1;
    char rsv2 = aHeader[0] & 0x20;  (void)rsv2;
    char rsv3 = aHeader[0] & 0x10;  (void)rsv3;
    char opcode = aHeader[0] & 0x0f;
    
    char mask = aHeader[1] & 0x80;
    uint64_t plen = aHeader[1] & 0x7f;
                                                                                    WSLO_("fin:%d, opcode:%d, mask:%d, plen:%d.", (int)fin, (int)opcode, (int)mask, (int)plen);
                                                                                    if(!fin) return -2;// NOT supported yet...
    
    char* pMask = NULL;
    
    if(plen == 0x7f){
        if(mask){
            if((n = recv(sock, aHeader+2, 8+4, MSG_WAITALL)) <= 0) return n;
            pMask = aHeader+(2+8);
        }else{
            if((n = recv(sock, aHeader+2, 8+0, MSG_WAITALL)) <= 0) return n;
        }
        
        #if BYTE_ORDER == BIG_ENDIAN
            plen = *(uint64_t*)&aHeader[2];
        #else
            ((char*)&plen)[0] = aHeader[2+7];
            ((char*)&plen)[1] = aHeader[2+6];
            ((char*)&plen)[2] = aHeader[2+5];
            ((char*)&plen)[3] = aHeader[2+4];
            ((char*)&plen)[4] = aHeader[2+3];
            ((char*)&plen)[5] = aHeader[2+2];
            ((char*)&plen)[6] = aHeader[2+1];
            ((char*)&plen)[7] = aHeader[2+0];
        #endif
    }else
    if(plen == 0x7e){
        if(mask){
            if((n = recv(sock, aHeader+2, 2+4, MSG_WAITALL)) <= 0) return n;
            pMask = aHeader+(2+2);
        }else{
            if((n = recv(sock, aHeader+2, 2+0, MSG_WAITALL)) <= 0) return n;
        }
        
        #if BYTE_ORDER == BIG_ENDIAN
            plen = *(uint16_t*)&aHeader[2];
        #else
            ((char*)&plen)[0] = aHeader[2+1];
            ((char*)&plen)[1] = aHeader[2+0];
        #endif
    }
    else{
        if(mask){
            if((n = recv(sock, aHeader+2, 0+4, MSG_WAITALL)) <= 0) return n;
            pMask = aHeader+(2+0);
        }else{
            //((n = recv(sock, aHeader+2, 0+0, MSG_WAITALL)) <= 0) return n;        // DO NOT exec because read size==0.
        }
    }
                                                                                    WSLO_("plen:%d.", (int)plen);
                                                                                    if(len < plen) return -2;// NOT enough buf size
    
    n = 0;
    if(plen){
        if((n = recv(sock, buf, plen, MSG_WAITALL)) <= 0) return n;
    }
    buf[n] = '\0';
    
    if(pMask){
        int i;
        for(i=0; i<n; ++i){
            buf[i] ^= pMask[i % 4];
        }
    }
                                                                                    WSLO_("payload:%s.", buf);
    
    if(popcode)
        *popcode = opcode;
    
    if(opcode & WS_OP_CTRL){
        if(opcode == WS_OP_CLOSE){
            ws_send(sock, NULL, 0, 0, NULL, WS_OP_CLOSE);
        }else
        if(opcode == WS_OP_PING){
            ws_send(sock, NULL, 0, 0, NULL, WS_OP_PONG);
        }
    }
    
    return n;
}









/*
 Copyright (c) 2011, Micael Hildenborg
 All rights reserved.

 Redistribution and use in source and binary forms, with or without
 modification, are permitted provided that the following conditions are met:
    * Redistributions of source code must retain the above copyright
      notice, this list of conditions and the following disclaimer.
    * Redistributions in binary form must reproduce the above copyright
      notice, this list of conditions and the following disclaimer in the
      documentation and/or other materials provided with the distribution.
    * Neither the name of Micael Hildenborg nor the
      names of its contributors may be used to endorse or promote products
      derived from this software without specific prior written permission.

 THIS SOFTWARE IS PROVIDED BY Micael Hildenborg ''AS IS'' AND ANY
 EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 DISCLAIMED. IN NO EVENT SHALL Micael Hildenborg BE LIABLE FOR ANY
 DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

/*
 Contributors:
 Gustav
 Several members in the gamedev.se forum.
 Gregory Petrosyan
 */

// Rotate an integer value to left.
inline static
unsigned int rol(const unsigned int value,
        const unsigned int steps)
{
    return ((value << steps) | (value >> (32 - steps)));
}

// Sets the first 16 integers in the buffert to zero.
// Used for clearing the W buffert.
inline static
void clearWBuffert(unsigned int* buffert)
{
    int pos;
    for (pos = 16; --pos >= 0;)
    {
        buffert[pos] = 0;
    }
}

static
void innerHash(unsigned int* result, unsigned int* w)
{
    unsigned int a = result[0];
    unsigned int b = result[1];
    unsigned int c = result[2];
    unsigned int d = result[3];
    unsigned int e = result[4];

    int round = 0;

    #define sha1macro(func,val) \
                { \
        const unsigned int t = rol(a, 5) + (func) + e + val + w[round]; \
                        e = d; \
                        d = c; \
                        c = rol(b, 30); \
                        b = a; \
                        a = t; \
                }

    while (round < 16)
    {
        sha1macro((b & c) | (~b & d), 0x5a827999)
        ++round;
    }
    while (round < 20)
    {
        w[round] = rol((w[round - 3] ^ w[round - 8] ^ w[round - 14] ^ w[round - 16]), 1);
        sha1macro((b & c) | (~b & d), 0x5a827999)
        ++round;
    }
    while (round < 40)
    {
        w[round] = rol((w[round - 3] ^ w[round - 8] ^ w[round - 14] ^ w[round - 16]), 1);
        sha1macro(b ^ c ^ d, 0x6ed9eba1)
        ++round;
    }
    while (round < 60)
    {
        w[round] = rol((w[round - 3] ^ w[round - 8] ^ w[round - 14] ^ w[round - 16]), 1);
        sha1macro((b & c) | (b & d) | (c & d), 0x8f1bbcdc)
        ++round;
    }
    while (round < 80)
    {
        w[round] = rol((w[round - 3] ^ w[round - 8] ^ w[round - 14] ^ w[round - 16]), 1);
        sha1macro(b ^ c ^ d, 0xca62c1d6)
        ++round;
    }

    #undef sha1macro

    result[0] += a;
    result[1] += b;
    result[2] += c;
    result[3] += d;
    result[4] += e;
}

static
void sha1_calc(const void* src, const int bytelength, unsigned char* hash)
{
    // Init the result array.
    unsigned int result[5] = { 0x67452301, 0xefcdab89, 0x98badcfe, 0x10325476, 0xc3d2e1f0 };

    // Cast the void src pointer to be the byte array we can work with.
    const unsigned char* sarray = (const unsigned char*) src;

    // The reusable round buffer
    unsigned int w[80];

    // Loop through all complete 64byte blocks.
    const int endOfFullBlocks = bytelength - 64;
    int endCurrentBlock;
    int currentBlock = 0;

    while (currentBlock <= endOfFullBlocks)
    {
        endCurrentBlock = currentBlock + 64;

        // Init the round buffer with the 64 byte block data.
        int roundPos;
        for (roundPos = 0; currentBlock < endCurrentBlock; currentBlock += 4)
        {
            // This line will swap endian on big endian and keep endian on little endian.
            w[roundPos++] = (unsigned int) sarray[currentBlock + 3]
                    | (((unsigned int) sarray[currentBlock + 2]) << 8)
                    | (((unsigned int) sarray[currentBlock + 1]) << 16)
                    | (((unsigned int) sarray[currentBlock]) << 24);
        }
        innerHash(result, w);
    }

    // Handle the last and not full 64 byte block if existing.
    endCurrentBlock = bytelength - currentBlock;
    clearWBuffert(w);
    int lastBlockBytes = 0;
    for (;lastBlockBytes < endCurrentBlock; ++lastBlockBytes)
    {
        w[lastBlockBytes >> 2] |= (unsigned int) sarray[lastBlockBytes + currentBlock] << ((3 - (lastBlockBytes & 3)) << 3);
    }
    w[lastBlockBytes >> 2] |= 0x80 << ((3 - (lastBlockBytes & 3)) << 3);
    if (endCurrentBlock >= 56)
    {
        innerHash(result, w);
        clearWBuffert(w);
    }
    w[15] = bytelength << 3;
    innerHash(result, w);

    // Store hash in result pointer, and make sure we get in in the correct order on both endian models.
    int hashByte;
    for (hashByte = 20; --hashByte >= 0;)
    {
        hash[hashByte] = (result[hashByte >> 2] >> (((3 - hashByte) & 0x3) << 3)) & 0xff;
    }
}









/*  Copyright (c) 2006-2008, Philip Busch <philip@0xe3.com>
 *  All rights reserved.
 *
 *  Redistribution and use in source and binary forms, with or without
 *  modification, are permitted provided that the following conditions are met:
 *
 *   - Redistributions of source code must retain the above copyright notice,
 *     this list of conditions and the following disclaimer.
 *   - Redistributions in binary form must reproduce the above copyright
 *     notice, this list of conditions and the following disclaimer in the
 *     documentation and/or other materials provided with the distribution.
 *
 *  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
 *  AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
 *  IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
 *  ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE
 *  LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
 *  CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
 *  SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
 *  INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
 *  CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
 *  ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
 *  POSSIBILITY OF SUCH DAMAGE.
 */

/**
 * @file
 * Base64 implementation.
 * @ingroup base64
 */

#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#define XX 100

/** @var base64_list
 *   A 64 character alphabet.
 *
 *   A 64-character subset of International Alphabet IA5, enabling
 *   6 bits to be represented per printable character.  (The proposed
 *   subset of characters is represented identically in IA5 and ASCII.)
 *   The character "=" signifies a special processing function used for
 *   padding within the printable encoding procedure.
 *
 *   \verbatim
    Value Encoding  Value Encoding  Value Encoding  Value Encoding
       0 A            17 R            34 i            51 z
       1 B            18 S            35 j            52 0
       2 C            19 T            36 k            53 1
       3 D            20 U            37 l            54 2
       4 E            21 V            38 m            55 3
       5 F            22 W            39 n            56 4
       6 G            23 X            40 o            57 5
       7 H            24 Y            41 p            58 6
       8 I            25 Z            42 q            59 7
       9 J            26 a            43 r            60 8
      10 K            27 b            44 s            61 9
      11 L            28 c            45 t            62 +
      12 M            29 d            46 u            63 /
      13 N            30 e            47 v
      14 O            31 f            48 w         (pad) =
      15 P            32 g            49 x
      16 Q            33 h            50 y
    \endverbatim
 */
static const char base64_list[] = \
        "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

/** Encode a minimal memory block. This function encodes a minimal memory area
 *  of three bytes into a printable base64-format sequence of four bytes.
 *  It is mainly used in more convenient functions, see below.
 *
 * @attention This function can't check if there's enough space at the memory
 *            memory location pointed to by \c out, so be careful.
 *
 * @param out pointer to destination
 * @param in pointer to source
 * @param len input size in bytes (between 0 and 3)
 * @returns nothing
 *
 * @ingroup base64
 */
static void base64_encode_block(unsigned char out[4], const unsigned char in[3], int len)
{
        out[0] = base64_list[ in[0] >> 2 ];
        out[1] = base64_list[ ((in[0] & 0x03) << 4) | ((in[1] & 0xf0) >> 4) ];
        out[2] = (unsigned char) (len > 1 ? base64_list[ ((in[1] & 0x0f) << 2) | ((in[2] & 0xc0) >> 6) ] : '=');
        out[3] = (unsigned char) (len > 2 ? base64_list[in[2] & 0x3f] : '=');
}

/** Encode an arbitrary size memory area. This function encodes the first
 *  \c len bytes of the contents of the memory area pointed to by \c in and
 *  stores the result in the memory area pointed to by \c out. The result will
 *  be null-terminated.
 *
 * @attention This function can't check if there's enough space at the memory
 *            memory location pointed to by \c out, so be careful.
 *
 * @param out pointer to destination
 * @param in pointer to source
 * @param len input size in bytes
 * @returns nothing
 *
 * @ingroup base64
 */
static void base64_encode_binary(char *out, const unsigned char *in, size_t len)
{
        int size;
        size_t i = 0;
        
        while(i < len) {
                size = (len-i < 4) ? len-i : 4;
                
                base64_encode_block((unsigned char *)out, in, size);

                out += 4;
                in  += 3;
                i   += 3;
        }

        *out = '\0';
}


