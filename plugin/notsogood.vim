" Send mail through Outlook

function! WinCallEncode (str)
    let a = substitute(a:str, " ", "\\\\%20", "g")
    let b = substitute(a, ",", "\\\\%2C", "g")
    return b
endfunction

function! OutlookMail (addresses, subject, body, attachment)
    if ! exists("g:notsogood_outlook_path")
        echoerr "g:notsogood_outlook_path not set!"
        return
    endif
    if ! file_readable(g:notsogood_outlook_path)
        echoerr "g:notsogood_outlook_path not a readable file!"
        return
    endif
    let outlook = substitute(g:notsogood_outlook_path, " ", "\\\\ ", "g")
    let cmd = outlook.' //c ipm.note'

    " refine the field data
    let addresses = join(a:addresses, ";")
    if len(a:subject) > 0
        let subject = "subject=".WinCallEncode(a:subject)
    else
        let subject = ""
    endif
    if len(a:body) > 0
        let body = "body=".WinCallEncode(a:body)
    else
        let body = ""
    endif

    " concatentate the field data
    let subbod = join([subject, body], "&")
    let msg = join([addresses, subbod], "?")

    " join any non-empty field data onto the message
    if len(msg) > 0
        let cmd = cmd.' //m "'.msg.'"'
    endif

    " handle optional attachment
    if a:attachment != ""
        let cmd = cmd.' //a "'.a:attachment.'"'
    endif

    echom "NotSoGood plugin executing the following command:"
    echom cmd

    exe "silent !".cmd
endfunction

function! EmailPeople ()
    echo "Whom shall we email?"
    for key in keys(g:notsogood_addrs)
        echo key.") ".split(g:notsogood_addrs[key], "|")[0]
    endfor
    let chosenOne = nr2char(getchar())
    if has_key(g:notsogood_addrs, chosenOne)
        call OutlookMail(
                    \ [split(g:notsogood_addrs[chosenOne], "|")[1]],
                    \ expand("%:p:t"),
                    \ "Sent from Vim.",
                    \ expand("%:p"))
    endif
    if exists("b:mailtemp")
        bd!
    endif
    redraw
endfunction

function! EmailSnippetTempFile ()
    let tempdir = fnamemodify(tempname(), ":p:h")
    let tempfile = fnamemodify(tempname(), ":p:t:r")
    let filename = expand("%:t:r")
    let extension = expand("%:e")
    if len(extension) > 0
        let extension = '.'.extension
    endif
    return tempdir."/snippet_".tempfile."_".filename.extension
endfunction

