<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Charset="utf-8"%>
<%
' Copyright (c) 2020, shenzhen haowen
'
' SPDX-License-Identifier: Apache-2.0
'
' Change Logs:
' Date           Author       		Notes
' 2020-12-05     812256@qq.com    	first version
%>
<%
private m_lonbits(30)
private m_l2power(30)
private k(63)
private const bits_to_a_byte = 8
private const bytes_to_a_word = 4
private const bits_to_a_word = 32
m_lonbits(0) = clng(1)
m_lonbits(1) = clng(3)
m_lonbits(2) = clng(7)
m_lonbits(3) = clng(15)
m_lonbits(4) = clng(31)
m_lonbits(5) = clng(63)
m_lonbits(6) = clng(127)
m_lonbits(7) = clng(255)
m_lonbits(8) = clng(511)
m_lonbits(9) = clng(1023)
m_lonbits(10) = clng(2047)
m_lonbits(11) = clng(4095)
m_lonbits(12) = clng(8191)
m_lonbits(13) = clng(16383)
m_lonbits(14) = clng(32767)
m_lonbits(15) = clng(65535)
m_lonbits(16) = clng(131071)
m_lonbits(17) = clng(262143)
m_lonbits(18) = clng(524287)
m_lonbits(19) = clng(1048575)
m_lonbits(20) = clng(2097151)
m_lonbits(21) = clng(4194303)
m_lonbits(22) = clng(8388607)
m_lonbits(23) = clng(16777215)
m_lonbits(24) = clng(33554431)
m_lonbits(25) = clng(67108863)
m_lonbits(26) = clng(134217727)
m_lonbits(27) = clng(268435455)
m_lonbits(28) = clng(536870911)
m_lonbits(29) = clng(1073741823)
m_lonbits(30) = clng(2147483647)
m_l2power(0) = clng(1)
m_l2power(1) = clng(2)
m_l2power(2) = clng(4)
m_l2power(3) = clng(8)
m_l2power(4) = clng(16)
m_l2power(5) = clng(32)
m_l2power(6) = clng(64)
m_l2power(7) = clng(128)
m_l2power(8) = clng(256)
m_l2power(9) = clng(512)
m_l2power(10) = clng(1024)
m_l2power(11) = clng(2048)
m_l2power(12) = clng(4096)
m_l2power(13) = clng(8192)
m_l2power(14) = clng(16384)
m_l2power(15) = clng(32768)
m_l2power(16) = clng(65536)
m_l2power(17) = clng(131072)
m_l2power(18) = clng(262144)
m_l2power(19) = clng(524288)
m_l2power(20) = clng(1048576)
m_l2power(21) = clng(2097152)
m_l2power(22) = clng(4194304)
m_l2power(23) = clng(8388608)
m_l2power(24) = clng(16777216)
m_l2power(25) = clng(33554432)
m_l2power(26) = clng(67108864)
m_l2power(27) = clng(134217728)
m_l2power(28) = clng(268435456)
m_l2power(29) = clng(536870912)
m_l2power(30) = clng(1073741824)
    
k(0) = &h428a2f98
k(1) = &h71374491
k(2) = &hb5c0fbcf
k(3) = &he9b5dba5
k(4) = &h3956c25b
k(5) = &h59f111f1
k(6) = &h923f82a4
k(7) = &hab1c5ed5
k(8) = &hd807aa98
k(9) = &h12835b01
k(10) = &h243185be
k(11) = &h550c7dc3
k(12) = &h72be5d74
k(13) = &h80deb1fe
k(14) = &h9bdc06a7
k(15) = &hc19bf174
k(16) = &he49b69c1
k(17) = &hefbe4786
k(18) = &hfc19dc6
k(19) = &h240ca1cc
k(20) = &h2de92c6f
k(21) = &h4a7484aa
k(22) = &h5cb0a9dc
k(23) = &h76f988da
k(24) = &h983e5152
k(25) = &ha831c66d
k(26) = &hb00327c8
k(27) = &hbf597fc7
k(28) = &hc6e00bf3
k(29) = &hd5a79147
k(30) = &h6ca6351
k(31) = &h14292967
k(32) = &h27b70a85
k(33) = &h2e1b2138
k(34) = &h4d2c6dfc
k(35) = &h53380d13
k(36) = &h650a7354
k(37) = &h766a0abb
k(38) = &h81c2c92e
k(39) = &h92722c85
k(40) = &ha2bfe8a1
k(41) = &ha81a664b
k(42) = &hc24b8b70
k(43) = &hc76c51a3
k(44) = &hd192e819
k(45) = &hd6990624
k(46) = &hf40e3585
k(47) = &h106aa070
k(48) = &h19a4c116
k(49) = &h1e376c08
k(50) = &h2748774c
k(51) = &h34b0bcb5
k(52) = &h391c0cb3
k(53) = &h4ed8aa4a
k(54) = &h5b9cca4f
k(55) = &h682e6ff3
k(56) = &h748f82ee
k(57) = &h78a5636f
k(58) = &h84c87814
k(59) = &h8cc70208
k(60) = &h90befffa
k(61) = &ha4506ceb
k(62) = &hbef9a3f7
k(63) = &hc67178f2
private function lshift(lvalue, ishiftbits)
    if ishiftbits = 0 then
        lshift = lvalue
        exit function
    elseif ishiftbits = 31 then
        if lvalue and 1 then
            lshift = &h80000000
        else
            lshift = 0
        end if
        exit function
    elseif ishiftbits < 0 or ishiftbits > 31 then
        err.raise 6
    end if
    
    if (lvalue and m_l2power(31 - ishiftbits)) then
        lshift = ((lvalue and m_lonbits(31 - (ishiftbits + 1))) * m_l2power(ishiftbits)) or &h80000000
    else
        lshift = ((lvalue and m_lonbits(31 - ishiftbits)) * m_l2power(ishiftbits))
    end if
end function
private function rshift(lvalue, ishiftbits)
    if ishiftbits = 0 then
        rshift = lvalue
        exit function
    elseif ishiftbits = 31 then
        if lvalue and &h80000000 then
            rshift = 1
        else
            rshift = 0
        end if
        exit function
    elseif ishiftbits < 0 or ishiftbits > 31 then
        err.raise 6
    end if
    
    rshift = (lvalue and &h7ffffffe) \ m_l2power(ishiftbits)
    
    if (lvalue and &h80000000) then
        rshift = (rshift or (&h40000000 \ m_l2power(ishiftbits - 1)))
    end if
end function
private function addunsigned(lx, ly)
    dim lx4
    dim ly4
    dim lx8
    dim ly8
    dim lresult
    lx8 = lx and &h80000000
    ly8 = ly and &h80000000
    lx4 = lx and &h40000000
    ly4 = ly and &h40000000
    lresult = (lx and &h3fffffff) + (ly and &h3fffffff)
    if lx4 and ly4 then
        lresult = lresult xor &h80000000 xor lx8 xor ly8
    elseif lx4 or ly4 then
        if lresult and &h40000000 then
            lresult = lresult xor &hc0000000 xor lx8 xor ly8
        else
            lresult = lresult xor &h40000000 xor lx8 xor ly8
        end if
    else
        lresult = lresult xor lx8 xor ly8
    end if
    addunsigned = lresult
end function
private function ch(x, y, z)
    ch = ((x and y) xor ((not x) and z))
end function
private function maj(x, y, z)
    maj = ((x and y) xor (x and z) xor (y and z))
end function
private function s(x, n)
    s = (rshift(x, (n and m_lonbits(4))) or lshift(x, (32 - (n and m_lonbits(4)))))
end function
private function r(x, n)
    r = rshift(x, cint(n and m_lonbits(4)))
end function
private function sigma0(x)
    sigma0 = (s(x, 2) xor s(x, 13) xor s(x, 22))
end function
private function sigma1(x)
    sigma1 = (s(x, 6) xor s(x, 11) xor s(x, 25))
end function
private function gamma0(x)
    gamma0 = (s(x, 7) xor s(x, 18) xor r(x, 3))
end function
private function gamma1(x)
    gamma1 = (s(x, 17) xor s(x, 19) xor r(x, 10))
end function
private function converttowordarray(smessage)
    dim lmessagelength
    dim lnumberofwords
    dim lwordarray()
    dim lbyteposition
    dim lbytecount
    dim lwordcount
    dim lbyte
    
    const modulus_bits = 512
    const congruent_bits = 448
    
    lmessagelength = len(smessage)
    
    lnumberofwords = (((lmessagelength + ((modulus_bits - congruent_bits) \ bits_to_a_byte)) \ (modulus_bits \ bits_to_a_byte)) + 1) * (modulus_bits \ bits_to_a_word)
    redim lwordarray(lnumberofwords - 1)
    
    lbyteposition = 0
    lbytecount = 0
    do until lbytecount >= lmessagelength
        lwordcount = lbytecount \ bytes_to_a_word
        
        lbyteposition = (3 - (lbytecount mod bytes_to_a_word)) * bits_to_a_byte
        
        lbyte = ascb(mid(smessage, lbytecount + 1, 1))
        
        lwordarray(lwordcount) = lwordarray(lwordcount) or lshift(lbyte, lbyteposition)
        lbytecount = lbytecount + 1
    loop
    lwordcount = lbytecount \ bytes_to_a_word
    lbyteposition = (3 - (lbytecount mod bytes_to_a_word)) * bits_to_a_byte
    lwordarray(lwordcount) = lwordarray(lwordcount) or lshift(&h80, lbyteposition)
    lwordarray(lnumberofwords - 1) = lshift(lmessagelength, 3)
    lwordarray(lnumberofwords - 2) = rshift(lmessagelength, 29)
    
    converttowordarray = lwordarray
end function
public function sha256(smessage)
    dim hash(7)
    dim m
    dim w(63)
    dim a
    dim b
    dim c
    dim d
    dim e
    dim f
    dim g
    dim h
    dim i
    dim j
    dim t1
    dim t2
    
    hash(0) = &h6a09e667
    hash(1) = &hbb67ae85
    hash(2) = &h3c6ef372
    hash(3) = &ha54ff53a
    hash(4) = &h510e527f
    hash(5) = &h9b05688c
    hash(6) = &h1f83d9ab
    hash(7) = &h5be0cd19
    
    m = converttowordarray(smessage)
    
    for i = 0 to ubound(m) step 16
        a = hash(0)
        b = hash(1)
        c = hash(2)
        d = hash(3)
        e = hash(4)
        f = hash(5)
        g = hash(6)
        h = hash(7)
        
        for j = 0 to 63
            if j < 16 then
                w(j) = m(j + i)
            else
                w(j) = addunsigned(addunsigned(addunsigned(gamma1(w(j - 2)), w(j - 7)), gamma0(w(j - 15))), w(j - 16))
            end if
                
            t1 = addunsigned(addunsigned(addunsigned(addunsigned(h, sigma1(e)), ch(e, f, g)), k(j)), w(j))
            t2 = addunsigned(sigma0(a), maj(a, b, c))
            
            h = g
            g = f
            f = e
            e = addunsigned(d, t1)
            d = c
            c = b
            b = a
            a = addunsigned(t1, t2)
        next
        
        hash(0) = addunsigned(a, hash(0))
        hash(1) = addunsigned(b, hash(1))
        hash(2) = addunsigned(c, hash(2))
        hash(3) = addunsigned(d, hash(3))
        hash(4) = addunsigned(e, hash(4))
        hash(5) = addunsigned(f, hash(5))
        hash(6) = addunsigned(g, hash(6))
        hash(7) = addunsigned(h, hash(7))
    next
    
    sha256 = lcase(right("00000000" & hex(hash(0)), 8) & right("00000000" & hex(hash(1)), 8) & right("00000000" & hex(hash(2)), 8) & right("00000000" & hex(hash(3)), 8) & right("00000000" & hex(hash(4)), 8) & right("00000000" & hex(hash(5)), 8) & right("00000000" & hex(hash(6)), 8) & right("00000000" & hex(hash(7)), 8))
end function
%>

<%
'把标准时间转换为UNIX时间戳
Function ToUnixTime(strTime)
    intTimeZone = 8
    If IsEmpty(strTime) or Not IsDate(strTime) Then strTime = Now
    If IsEmpty(intTimeZone) or Not isNumeric(intTimeZone) Then intTimeZone = 0
    ToUnixTime = DateAdd("h",-intTimeZone,strTime)
    ToUnixTime = DateDiff("s","1970-01-01 00:00:00", ToUnixTime)
End Function
 
'把UNIX时间戳转换为标准时间
Function FromUnixTime(intTime)
    intTimeZone = 8
    If IsEmpty(intTime) or Not IsNumeric(intTime) Then
        FromUnixTime = Now()
        Exit Function
    End If         
    If IsEmpty(intTime) or Not IsNumeric(intTimeZone) Then intTimeZone = 0
    FromUnixTime = DateAdd("s", intTime, "1970-01-01 00:00:00")
    FromUnixTime = DateAdd("h", intTimeZone, FromUnixTime)
End Function
%>