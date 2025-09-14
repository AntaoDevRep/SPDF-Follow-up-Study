
text.font.size <- 10

## normal 
m.normal.theme <- theme(text = element_text(size = text.font.size,  family="sans", color = "black"),
                 plot.title = element_text(size = text.font.size*1.2, face="bold"),
                 axis.title = element_text(size = text.font.size, color = "black", face="bold"),
                 axis.text = element_text(size = text.font.size, color = "black"),
                 strip.text = element_text(size = text.font.size, color = "black", face="bold"),
                 legend.text = element_text(size = text.font.size, color = "black"),
                 plot.caption = element_text(size = text.font.size, color = "black"),
                 panel.grid.major = element_blank(), # add horizontal grid
                 panel.grid.minor = element_blank(), # remove grid
                 panel.background = element_blank(), # remove background
                 axis.line = element_line(colour = "black"), # draw axis line in black
                 legend.position = "right")

## without legend 
m.no.legend.theme <- theme(text = element_text(size = text.font.size,  family="sans", color = "black"),
                        plot.title = element_text(size = text.font.size*1.2, face="bold"),
                        axis.title = element_text(size = text.font.size, color = "black", face="bold"),
                        axis.text = element_text(size = text.font.size, color = "black"),
                        strip.text = element_text(size = text.font.size, color = "black", face="bold"),
                        legend.text = element_text(size = text.font.size, color = "black"),
                        plot.caption = element_text(size = text.font.size, color = "black"),
                        panel.grid.major = element_blank(), # add horizontal grid
                        panel.grid.minor = element_blank(), # remove grid
                        panel.background = element_blank(), # remove background
                        axis.line = element_line(colour = "black"), # draw axis line in black
                        legend.position = "none")

## without legend 
m.top.legend.theme <- theme(text = element_text(size = text.font.size,  family="sans", color = "black"),
                           plot.title = element_text(size = text.font.size*1.2, face="bold"),
                           axis.title = element_text(size = text.font.size, color = "black", face="bold"),
                           axis.text = element_text(size = text.font.size, color = "black"),
                           strip.text = element_text(size = text.font.size, color = "black", face="bold"),
                           legend.text = element_text(size = text.font.size, color = "black"),
                           plot.caption = element_text(size = text.font.size, color = "black"),
                           panel.grid.major = element_blank(), # add horizontal grid
                           panel.grid.minor = element_blank(), # remove grid
                           panel.background = element_blank(), # remove background
                           axis.line = element_line(colour = "black"), # draw axis line in black
                           legend.position = "top")


survival.plot.theme <- theme(text = element_text(size = text.font.size,  family="sans", color = "black"), 
                             plot.title = element_text(size = text.font.size*1.2, face="bold"),
                             plot.caption = element_text(size = text.font.size, color = "black"),
                             axis.title = element_text(size = text.font.size, color = "black", face="bold"),
                             axis.text = element_text(size = text.font.size, color = "black"),
                             strip.text = element_text(size = text.font.size, color = "black", face="bold"),
                             legend.text = element_text(size = text.font.size, color = "black"),
                             panel.grid.major = element_blank(), # remove vertical grid
                             panel.grid.minor = element_blank(), # remove grid
                             panel.background = element_blank(), # remove background
                             axis.line = element_line(colour = "black"), # draw axis line in black
                             legend.position = "none")

## without legend 
m.no.legend.grey.bg.theme <- theme(text = element_text(size = text.font.size,  family="sans", color = "black"),
                           plot.title = element_text(size = text.font.size*1.2, face="bold"),
                           axis.title = element_text(size = text.font.size, color = "black", face="bold"),
                           axis.text = element_text(size = text.font.size, color = "black"),
                           strip.text = element_text(size = text.font.size, color = "black", face="bold"),
                           legend.text = element_text(size = text.font.size, color = "black"),
                           plot.caption = element_text(size = text.font.size, color = "black"),
                           panel.grid.major = element_line(color = "white"), # add horizontal grid
                           panel.grid.minor = element_line(color = "white"), 
                           panel.background = element_rect(fill = "#eeeeee"), # grey background
                           axis.line = element_line(colour = "black"), # draw axis line in black
                           legend.position = "none")



## normal 
m.normal.grey.grid.theme <- theme(text = element_text(size = text.font.size,  family="sans", color = "black"),
                        plot.title = element_text(size = text.font.size*1.2, face="bold"),
                        axis.title = element_text(size = text.font.size, color = "black", face="bold"),
                        axis.text = element_text(size = text.font.size, color = "black"),
                        strip.text = element_text(size = text.font.size, color = "black", face="bold"),
                        legend.text = element_text(size = text.font.size, color = "black"),
                        plot.caption = element_text(size = text.font.size, color = "black"),
                        panel.grid.major = element_line(color = "#dddddd"), # add horizontal grid
                        panel.grid.minor = element_blank(), # remove grid
                        panel.background = element_blank(), # remove background
                        axis.line = element_line(colour = "black"), # draw axis line in black
                        legend.position = "right")