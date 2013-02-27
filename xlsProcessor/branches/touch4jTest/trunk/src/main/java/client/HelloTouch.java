package client;

import com.emitrom.touch4j.client.core.TouchEntryPoint;
import com.emitrom.touch4j.client.ui.ViewPort;

/**
 * Created with IntelliJ IDEA. User: LiHa Date: 2/27/13
 */
public class HelloTouch extends TouchEntryPoint
{
    private static final HelloTouch INSTANCE = new HelloTouch();
    private Content content;

    public static HelloTouch getInstance(){
        return INSTANCE;
    }

    public void onLoad()
    {
//        Panel panel = new Panel();
//        Button hello = new Button("Say Hello");
//        hello.addTapHandler(new TapHandler() {
//            @Override
//            public void onTap(Button button, EventObject event)
//            {
//                MessageBox.alert("Hello World");
//            }
//        });
//        panel.add(hello);
//        hello.setCentered(true);
//        ViewPort.get().add(panel);


        HelloTouch.getInstance().setContent(new MyGoogleMap());
    }

    public void setContent(Content content){
        this.content = content;
        ViewPort.get().clear();
        ViewPort.get().add(content);
    }
}
