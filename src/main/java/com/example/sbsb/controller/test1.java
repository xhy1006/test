package com.example.sbsb.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;


@Controller
public class test1 {

    @RequestMapping("/test")
    public ModelAndView getIndex() {
        ModelAndView mo = new ModelAndView();
        mo.addObject("mvc","hello");
        mo.setViewName("index");

        return mo;
    }

}
