package client;

import com.emitrom.touch4j.client.ui.Panel;
import com.google.gwt.core.client.GWT;
import com.google.gwt.uibinder.client.UiBinder;

/**
 * Created with IntelliJ IDEA. User: LiHa Date: 2/27/13
 */
public class SecondView extends Content
{
    interface SecondViewUiBinder extends UiBinder<Panel, SecondView>
    {
    }

    private static SecondViewUiBinder ourUiBinder = GWT.create(SecondViewUiBinder.class);

    public SecondView()
    {
        Panel rootElement = ourUiBinder.createAndBindUi(this);
        initWidget(rootElement);
    }
}