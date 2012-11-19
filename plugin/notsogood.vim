" Send mail through Outlook
let g:outlook_path="C:/Progra~2/Micros~2/Office12/Outlook.exe"

function! OutlookMail (addresses, subject, body, attachment)
    let cmd=g:outlook_path.' /c ipm.note'

    " refine the field data
    let addrs=join(a:addresses, ";")
    if len(a:subject) > 0
        let subject = "subject=".a:subject
    else
        let subject = ""
    endif
    if len(a:body) > 0
        let body = "body=".a:body
    else
        let body = ""
    endif

    " concatentate the field data
    let msg=join([addrs, subject, body], "&")

    " join any non-empty field data onto the message
    if len(msg) > 0
        let cmd = cmd.' /m "'.msg.'"'
    endif

    " handle optional attachment
    if a:attachment != ""
        let cmd = cmd.' /a "'.a:attachment.'"'
    endif

    "echo cmd
    exe "silent !".cmd
endfunction

function! EmailPeople ()
    echo "Whom shall we email?"
    for key in keys(g:notsogood_addrs)
        echo key.") ".split(g:notsogood_addrs[key], "|")[0]
    endfor
    let chosenOne=nr2char(getchar())
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

nnoremap <Leader>E :call EmailPeople()<CR>
vnoremap <expr> <Leader>E "y:new ".EmailSnippetTempFile()."<CR>:let b:mailtemp=1<CR>VGP:w<CR>:call EmailPeople()<CR>"

