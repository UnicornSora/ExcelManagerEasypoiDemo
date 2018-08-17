package com.caterpillar.exceldatamanager.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

@Controller
public class IndexController {

    /**
     * 本地访问内容地址 ：http://localhost:8080/index
     *
     * @return
     */
    @RequestMapping("/index")
    public String helloHtml() {
        return "/index";
    }
}
