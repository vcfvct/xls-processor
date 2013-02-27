package client;

import com.emitrom.touch4j.client.core.EventObject;
import com.emitrom.touch4j.client.core.handlers.button.TapHandler;
import com.emitrom.touch4j.client.ui.Button;
import com.emitrom.touch4j.client.ui.MessageBox;
import com.emitrom.touch4j.client.ui.Panel;
import com.google.gwt.core.client.GWT;
import com.google.gwt.uibinder.client.UiBinder;
import com.google.gwt.uibinder.client.UiField;

/**
 * Created with IntelliJ IDEA. User: LiHa Date: 2/27/13
 */
public class MyGoogleMap extends Content
{
    interface googleMapUiBinder extends UiBinder<Panel, MyGoogleMap>
    {
    }

    private static googleMapUiBinder ourUiBinder = GWT.create(googleMapUiBinder.class);

    @UiField
    Button infoButton;
    @UiField
    Button backButton;

    public MyGoogleMap()
    {
        Panel rootElement = ourUiBinder.createAndBindUi(this);

        infoButton.addTapHandler(new TapHandler()
        {
            @Override
            public void onTap(Button button, EventObject event)
            {
                MessageBox.alert("Welcome Here");
            }
        });

        backButton.addTapHandler(new TapHandler()
        {
            @Override
            public void onTap(Button button, EventObject eventObject)
            {
                 HelloTouch.getInstance().setContent(new SecondView());
            }
        });

        initWidget(rootElement);
    }
}