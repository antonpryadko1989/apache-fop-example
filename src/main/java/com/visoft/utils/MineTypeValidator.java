package com.visoft.utils;

import net.sf.jmimemagic.Magic;
import net.sf.jmimemagic.MagicException;
import net.sf.jmimemagic.MagicMatchNotFoundException;
import net.sf.jmimemagic.MagicParseException;

public class MineTypeValidator {

    public static String getMineTypeFromByteArray(byte[] data) throws MagicException, MagicMatchNotFoundException, MagicParseException {
        return Magic.getMagicMatch(data, true).getMimeType();
    }
}
